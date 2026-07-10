import * as React from 'react';
import type {
  IAccessGrants,
  IAccountSettings,
  IEmploymentInfo,
  IIdentityInfo,
  ILicenseSelection,
  IPersonalInfo
} from '../models';
import { EMPTY_ACCESS_GRANTS } from '../models';

export interface IWizardDraft {
  personal: IPersonalInfo;
  employment: IEmploymentInfo;
  identity: IIdentityInfo | null;
  accountSettings: IAccountSettings;
  licenses: ILicenseSelection[];
  access: IAccessGrants;
  expirationReviewDays: number | null;
  /** Set via the "Clone access from" picker — makes this a Clone job. */
  cloneSourceUserId?: string;
  cloneSourceDisplayName?: string;
  /**
   * Snapshotted from the applied template's approverGroupId (null if the
   * template set none, or none was applied) — gates who can approve this
   * job once submitted. See JobDetailDrawer's showApprove logic.
   */
  approverGroupId: string | null;
}

export interface IWizardState {
  step: number;
  draft: IWizardDraft;
  /**
   * Tracks "the operator has entered something or advanced past step 1",
   * maintained incrementally by the reducer instead of derived by deep-
   * comparing the draft on every render (isDraftDirty used to JSON.stringify
   * it on every keystroke, in every consumer). Sticky by design: once set it
   * stays true until 'reset', even if edits are undone back to blank — the
   * same trade-off most form "dirty" trackers make.
   */
  dirty: boolean;
}

/** Personal, Employment, Identity, Account, Licenses, Access, Review. */
export const WIZARD_STEP_COUNT: number = 7;

export const initialWizardState: IWizardState = {
  step: 0,
  dirty: false,
  draft: {
    personal: {
      firstName: '',
      lastName: '',
      displayName: '',
      employeeId: '',
      mobilePhone: '',
      personalEmail: '',
      photoDataUrl: undefined
    },
    employment: {
      jobTitle: '',
      department: '',
      companyName: '',
      officeLocation: '',
      country: '',
      managerId: undefined,
      managerDisplayName: undefined,
      managerUpn: undefined,
      employeeType: 'Employee',
      hireDate: ''
    },
    identity: null,
    accountSettings: {
      usageLocation: 'SE',
      accountEnabled: true,
      credentialMode: 'tap',
      forceChangePassword: true
    },
    licenses: [],
    access: { ...EMPTY_ACCESS_GRANTS },
    expirationReviewDays: null,
    approverGroupId: null
  }
};

export type WizardAction =
  | { type: 'savePersonal'; personal: IPersonalInfo }
  | { type: 'saveEmployment'; employment: IEmploymentInfo }
  | { type: 'saveIdentity'; identity: IIdentityInfo | null }
  | { type: 'saveAccountSettings'; accountSettings: IAccountSettings }
  | { type: 'toggleLicense'; license: ILicenseSelection }
  | { type: 'saveAccess'; access: IAccessGrants; expirationReviewDays?: number | null }
  | { type: 'setCloneSource'; userId: string | undefined; displayName: string | undefined }
  | {
      /** Seeds employment/account/license/access defaults from a department template. */
      type: 'applyTemplate';
      department: string;
      usageLocation?: string;
      licenses: ILicenseSelection[];
      access?: Partial<IAccessGrants>;
      expirationReviewDays?: number | null;
      approverGroupId?: string | null;
    }
  | { type: 'next' }
  | { type: 'back' }
  | { type: 'goto'; step: number }
  | { type: 'reset' };

export function wizardReducer(state: IWizardState, action: WizardAction): IWizardState {
  switch (action.type) {
    case 'savePersonal':
      return { ...state, dirty: true, draft: { ...state.draft, personal: action.personal } };
    case 'saveEmployment':
      return { ...state, dirty: true, draft: { ...state.draft, employment: action.employment } };
    case 'saveIdentity':
      return { ...state, dirty: true, draft: { ...state.draft, identity: action.identity } };
    case 'saveAccountSettings':
      return {
        ...state,
        dirty: true,
        draft: { ...state.draft, accountSettings: action.accountSettings }
      };
    case 'toggleLicense': {
      const exists: boolean = state.draft.licenses.some((l) => l.skuId === action.license.skuId);
      const licenses: ILicenseSelection[] = exists
        ? state.draft.licenses.filter((l) => l.skuId !== action.license.skuId)
        : [...state.draft.licenses, action.license];
      return { ...state, dirty: true, draft: { ...state.draft, licenses } };
    }
    case 'saveAccess':
      return {
        ...state,
        dirty: true,
        draft: {
          ...state.draft,
          access: action.access,
          expirationReviewDays:
            action.expirationReviewDays !== undefined
              ? action.expirationReviewDays
              : state.draft.expirationReviewDays
        }
      };
    case 'setCloneSource':
      return {
        ...state,
        dirty: true,
        draft: {
          ...state.draft,
          cloneSourceUserId: action.userId,
          cloneSourceDisplayName: action.displayName
        }
      };
    case 'applyTemplate': {
      // Merge template licenses over the user's existing selections — any
      // template license not already selected is added. Existing manual
      // selections are preserved (no silent data loss). Access grants merge
      // the same way (union, no duplicates); expirationReviewDays overwrites
      // since it's a single tenant-set policy value, not a per-item list.
      const existingSkus: Set<string> = new Set(state.draft.licenses.map((l) => l.skuId));
      const merged: ILicenseSelection[] = [...state.draft.licenses];
      for (const tpl of action.licenses) {
        if (!existingSkus.has(tpl.skuId)) {
          merged.push(tpl);
          existingSkus.add(tpl.skuId);
        }
      }
      const mergedAccess: IAccessGrants = mergeAccessGrants(state.draft.access, action.access);
      return {
        ...state,
        dirty: true,
        draft: {
          ...state.draft,
          employment: action.department
            ? { ...state.draft.employment, department: action.department }
            : state.draft.employment,
          accountSettings: action.usageLocation
            ? { ...state.draft.accountSettings, usageLocation: action.usageLocation }
            : state.draft.accountSettings,
          licenses: merged,
          access: mergedAccess,
          expirationReviewDays:
            action.expirationReviewDays !== undefined
              ? action.expirationReviewDays
              : state.draft.expirationReviewDays,
          // Replace, not merge — approval routing is a single authoritative
          // value from whichever template was applied most recently, unlike
          // access grants which accumulate.
          approverGroupId: action.approverGroupId ?? null
        }
      };
    }
    case 'next': {
      const step: number = Math.min(state.step + 1, WIZARD_STEP_COUNT - 1);
      return { ...state, step, dirty: state.dirty || step > 0 };
    }
    case 'back':
      return { ...state, step: Math.max(state.step - 1, 0) };
    case 'goto': {
      const step: number = Math.min(Math.max(action.step, 0), WIZARD_STEP_COUNT - 1);
      return { ...state, step, dirty: state.dirty || step > 0 };
    }
    case 'reset':
      return { ...initialWizardState };
    default:
      return state;
  }
}

function mergeAccessGrants(current: IAccessGrants, incoming?: Partial<IAccessGrants>): IAccessGrants {
  if (!incoming) {
    return current;
  }
  const unionStrings = (a: string[], b?: string[]): string[] => Array.from(new Set([...a, ...(b ?? [])]));
  return {
    securityGroups: unionStrings(current.securityGroups, incoming.securityGroups),
    m365Groups: unionStrings(current.m365Groups, incoming.m365Groups),
    teams: unionByKey(current.teams, incoming.teams, (t) => t.teamId),
    sharePointSites: unionByKey(current.sharePointSites, incoming.sharePointSites, (s) => s.siteUrl),
    applications: unionStrings(current.applications, incoming.applications)
  };
}

function unionByKey<T>(current: T[], incoming: T[] | undefined, keyOf: (item: T) => string): T[] {
  if (!incoming || incoming.length === 0) {
    return current;
  }
  const existingKeys: Set<string> = new Set(current.map(keyOf));
  const merged: T[] = [...current];
  for (const item of incoming) {
    if (!existingKeys.has(keyOf(item))) {
      merged.push(item);
      existingKeys.add(keyOf(item));
    }
  }
  return merged;
}

/** True once the operator has entered anything or advanced past step 1. */
export function isDraftDirty(state: IWizardState): boolean {
  return state.dirty;
}

interface IWizardContextValue {
  state: IWizardState;
  dispatch: React.Dispatch<WizardAction>;
}

const WizardContext: React.Context<IWizardContextValue | undefined> = React.createContext<
  IWizardContextValue | undefined
>(undefined);

export const WizardProvider: React.FC<{ children?: React.ReactNode }> = ({ children }) => {
  const [state, dispatch] = React.useReducer(wizardReducer, initialWizardState);
  const value: IWizardContextValue = React.useMemo(() => ({ state, dispatch }), [state]);
  return <WizardContext.Provider value={value}>{children}</WizardContext.Provider>;
};

export function useWizard(): IWizardContextValue {
  const value: IWizardContextValue | undefined = React.useContext(WizardContext);
  if (!value) {
    throw new Error('useWizard must be used inside a WizardProvider');
  }
  return value;
}
