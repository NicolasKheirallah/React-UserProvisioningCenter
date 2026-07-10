import { spfi, SPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/items';
import { AddFieldOptions } from '@pnp/sp/fields';
import type { IList } from '@pnp/sp/lists';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { UPC_LIST_DEFINITIONS } from './listSchemas';
import type { IUpcFieldDefinition, IUpcListDefinition } from './listSchemas';
import { UPC_SEED_DEFINITIONS, type ISeedDefinition } from './listSeedItems';
import { escapeODataLiteral } from '../util/odata';

export interface IProvisioningProgress {
  message: string;
}

export interface IProvisioningResult {
  createdLists: string[];
  existingLists: string[];
  createdFields: number;
  /** Number of seed list items inserted (settings row, role rows, …). */
  createdItems: number;
}

const GENERIC_LIST_TEMPLATE: number = 100;

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function fieldSchemaXml(field: IUpcFieldDefinition): string {
  const common: string =
    `Name="${field.name}" StaticName="${field.name}" ` +
    `DisplayName="${escapeXml(field.displayName)}"` +
    (field.required ? ' Required="TRUE"' : '');
  switch (field.type) {
    case 'Note':
      return `<Field Type="Note" ${common} NumLines="6" />`;
    case 'DateTime':
      return `<Field Type="DateTime" ${common} Format="DateTime" />`;
    case 'User':
      return `<Field Type="User" ${common} UserSelectionMode="PeopleOnly" />`;
    case 'Choice': {
      const choices: string = (field.choices ?? [])
        .map((c) => `<CHOICE>${escapeXml(c)}</CHOICE>`)
        .join('');
      return `<Field Type="Choice" ${common} Format="Dropdown"><CHOICES>${choices}</CHOICES></Field>`;
    }
    default:
      return `<Field Type="${field.type}" ${common} />`;
  }
}

/**
 * Client-side alternative to provisioning-assets/lists.ps1: creates the
 * UPC_* lists from the web part's property pane. Idempotent — existing
 * lists/fields are left untouched. Requires the signed-in user to hold
 * Manage Lists rights on the host site (site owner); this is SharePoint
 * permission territory, independent of the Graph/Entra model.
 */
export class ListProvisioningService {
  private readonly _sp: SPFI;

  public constructor(context: WebPartContext);
  public constructor(sp: SPFI);
  public constructor(contextOrSp: WebPartContext | SPFI) {
    // Share the SPFI root with SharePointDataService when callers pass the
    // same instance; otherwise create a standalone one (property-pane path).
    this._sp =
      contextOrSp && typeof contextOrSp === 'object' && 'web' in contextOrSp
        ? (contextOrSp as SPFI)
        : spfi().using(SPFx(contextOrSp as WebPartContext));
  }

  public async ensureAllLists(
    onProgress?: (progress: IProvisioningProgress) => void
  ): Promise<IProvisioningResult> {
    const result: IProvisioningResult = {
      createdLists: [],
      existingLists: [],
      createdFields: 0,
      createdItems: 0
    };
    for (const definition of UPC_LIST_DEFINITIONS) {
      onProgress?.({ message: definition.title });
      const created: number = await this._ensureList(definition, result);
      result.createdFields += created;
    }
    // Seed default list items (UPC_Settings row, UPC_Roles rows) after all
    // lists and fields exist. Idempotent — only inserts what is still missing.
    for (const seed of UPC_SEED_DEFINITIONS) {
      onProgress?.({ message: seed.listTitle });
      result.createdItems += await this._ensureSeedItems(seed);
    }
    return result;
  }

  private async _ensureList(
    definition: IUpcListDefinition,
    result: IProvisioningResult
  ): Promise<number> {
    const ensure = await this._sp.web.lists.ensure(
      definition.title,
      'User Provisioning Center',
      GENERIC_LIST_TEMPLATE
    );
    if (ensure.created) {
      result.createdLists.push(definition.title);
    } else {
      result.existingLists.push(definition.title);
    }
    const list: IList = ensure.list;

    const existingFields: { InternalName: string; Title: string; CanBeDeleted: boolean }[] =
      await list.fields.select('InternalName', 'Title', 'CanBeDeleted')();
    const existingNames: Set<string> = new Set(existingFields.map((f) => f.InternalName));

    let createdFields: number = 0;
    for (const field of definition.fields) {
      if (existingNames.has(field.name)) {
        continue;
      }
      // Repair pass: an earlier version created fields without the
      // internal-name hint, so SharePoint derived internal names from the
      // display name (e.g. Job_x0020_Type). Remove such a stray before
      // creating the correctly named field.
      const mangled = existingFields.filter(
        (f) =>
          f.Title === field.displayName &&
          f.InternalName !== field.name &&
          f.InternalName.indexOf('_x0') !== -1 &&
          f.CanBeDeleted
      )[0];
      if (mangled) {
        await list.fields.getByInternalNameOrTitle(mangled.InternalName).delete();
      }
      await list.fields.createFieldAsXml({
        SchemaXml: fieldSchemaXml(field),
        // Without AddFieldInternalNameHint SharePoint ignores the Name
        // attribute and mangles the internal name from DisplayName.
        Options:
          AddFieldOptions.AddFieldInternalNameHint | AddFieldOptions.AddFieldToDefaultView
      });
      createdFields++;
    }

    if (definition.auditSecurity) {
      // Members may only edit their own items; versioning keeps history.
      // Closest list-level approximation of create-only (see lists.ps1 note).
      // Checked and repaired on every run — not just at first creation — so
      // re-running "Provision Lists" actually restores these settings if an
      // admin or governance tool has since reset them on an existing list.
      const current: { ReadSecurity: number; WriteSecurity: number; EnableVersioning: boolean } =
        await list.select('ReadSecurity', 'WriteSecurity', 'EnableVersioning')();
      if (
        current.ReadSecurity !== 1 ||
        current.WriteSecurity !== 2 ||
        !current.EnableVersioning
      ) {
        await list.update({ ReadSecurity: 1, WriteSecurity: 2, EnableVersioning: true });
      }
    }
    return createdFields;
  }

  /**
   * Inserts seed items when missing. Detects existing items by the
   * `identityField` (Title) so re-running never duplicates them. If a row
   * already exists but its payload is empty (e.g. MemberGroupId blank on a
   * pre-existing UPC_Roles row), the row is left untouched — the admin may
   * have intentionally cleared it.
   */
  private async _ensureSeedItems(seed: ISeedDefinition): Promise<number> {
    const list = this._sp.web.lists.getByTitle(seed.listTitle);
    let created: number = 0;
    for (const item of seed.items) {
      const identityValue: string = String(item[seed.identityField] ?? '');
      const existing: { Id: number }[] = await list.items
        .select('Id')
        .filter(`${seed.identityField} eq '${escapeODataLiteral(identityValue)}'`)
        .top(1)();
      if (existing.length > 0) {
        continue;
      }
      await list.items.add(item);
      created++;
    }
    return created;
  }
}
