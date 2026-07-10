import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { Version } from '@microsoft/sp-core-library';
import {
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  ThemeProvider,
  ThemeChangedEventArgs,
  type IReadonlyTheme
} from '@microsoft/sp-component-base';
import * as strings from 'ProvisioningCenterWebPartStrings';
import { App } from '../../components/App';
import type { IAppProps } from '../../components/App';
import { createServices } from '../../services/createServices';
import type { IServices } from '../../services/createServices';
import { createAppTheme, type TeamsAppTheme } from '../../theme/createAppTheme';
import { ListProvisioningService } from '../../services/provisioning/ListProvisioningService';

// Behavioral settings live in the UPC_Settings list (in-app Settings tab,
// manageSettings permission) — the property pane only hosts provisioning.
export interface IProvisioningCenterWebPartProps {}

export default class ProvisioningCenterWebPart extends BaseClientSideWebPart<IProvisioningCenterWebPartProps> {
  private _services: IServices | undefined;
  private _themeProvider: ThemeProvider | undefined;
  private _themeVariant: IReadonlyTheme | undefined;
  private _provisioning: boolean = false;
  private _provisionStatus: string = '';
  // React 18 root, created once per web part instance and reused across
  // SPFx's repeated render() calls (theme changes, property pane refresh).
  private _root: Root | undefined;

  protected async onInit(): Promise<void> {
    await super.onInit();
    // Section theme variants + site theme (supportsThemeVariants in the manifest).
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._themeVariant = this._themeProvider.tryGetTheme();
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChanged);
    // Composition root: MSGraphClientV3 (delegated) + PnPJS + engine.
    // A failure here (missing Graph consent, no SPFx context) must surface
    // clearly instead of leaving the web part blank.
    try {
      this._services = await createServices(this.context);
      this._provisionStatus = strings.ProvisionStatusIdle;
    } catch (err) {
      const message: string = err instanceof Error ? err.message : String(err);
      this._provisionStatus = `${strings.ProvisionStatusError} ${message}`.trim();
      console.error('[UPC] service composition failed', err);
    }
  }

  private _handleThemeChanged = (args: ThemeChangedEventArgs): void => {
    this._themeVariant = args.theme;
    this.render();
  };

  /** Teams host theme; SPFx exposes TeamsJS v1 (theme) or v2 (app.theme) context shapes. */
  private _getTeamsTheme(): TeamsAppTheme | undefined {
    const teams = this.context.sdks.microsoftTeams;
    if (!teams) {
      return undefined;
    }
    const ctx = teams.context as { theme?: string; app?: { theme?: string } } | undefined;
    const theme: string | undefined = ctx?.app?.theme ?? ctx?.theme;
    if (theme === 'dark' || theme === 'contrast') {
      return theme;
    }
    return theme ? 'default' : undefined;
  }

  public render(): void {
    if (!this._services) {
      return;
    }
    const element: React.ReactElement<IAppProps> = React.createElement(App, {
      services: this._services,
      theme: createAppTheme(this._themeVariant, this._getTeamsTheme()),
      dir: document.documentElement.dir === 'rtl' ? 'rtl' : 'ltr',
      instanceId: this.context.instanceId
    });
    if (!this._root) {
      this._root = createRoot(this.domElement);
    }
    this._root.render(element);
  }

  protected onDispose(): void {
    // Unmount defers its DOM cleanup to a microtask; React 18 warns if a
    // synchronous unmount races that, so don't call render() after this.
    this._root?.unmount();
    this._root = undefined;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * Property-pane alternative to provisioning-assets/lists.ps1. Requires
   * Manage Lists rights on the host site; idempotent, so re-running after a
   * partial failure only creates what is still missing.
   */
  private _onProvisionListsClick = (): void => {
    void this._provisionLists();
  };

  private async _provisionLists(): Promise<void> {
    if (this._provisioning) {
      return;
    }
    this._provisioning = true;
    this._provisionStatus = strings.ProvisionStatusRunning;
    this.context.propertyPane.refresh();
    try {
      const service: ListProvisioningService = new ListProvisioningService(this.context);
      const result = await service.ensureAllLists((progress) => {
        this._provisionStatus = `${strings.ProvisionStatusRunning} ${progress.message}`;
        this.context.propertyPane.refresh();
      });
      this._provisionStatus =
        `${strings.ProvisionStatusDone} ` +
        `${strings.ProvisionCreatedLabel}: ${result.createdLists.length}, ` +
        `${strings.ProvisionExistingLabel}: ${result.existingLists.length}, ` +
        `${strings.ProvisionFieldsLabel}: ${result.createdFields}, ` +
        `${strings.ProvisionItemsLabel}: ${result.createdItems}`;
      this.render();
    } catch (err) {
      const message: string = err instanceof Error ? err.message : '';
      this._provisionStatus = `${strings.ProvisionStatusError} ${message}`.trim();
    } finally {
      this._provisioning = false;
      this.context.propertyPane.refresh();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.ProvisionGroupName,
              groupFields: [
                PropertyPaneLabel('provisionIntro', {
                  text: strings.ProvisionIntro
                }),
                PropertyPaneButton('provisionLists', {
                  text: strings.ProvisionButtonLabel,
                  buttonType: PropertyPaneButtonType.Primary,
                  disabled: this._provisioning,
                  onClick: this._onProvisionListsClick
                }),
                PropertyPaneLabel('provisionStatus', {
                  text: this._provisionStatus
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
