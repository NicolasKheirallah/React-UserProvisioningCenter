import * as React from 'react';
import {
  Badge,
  Caption1,
  FluentProvider,
  IdPrefixProvider,
  ProgressBar,
  Spinner,
  Tab,
  TabList,
  Title3,
  makeStyles,
  tokens,
  webLightTheme,
  type Theme
} from '@fluentui/react-components';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import * as strings from 'UpcStrings';
import type { IServices } from '../services/createServices';
import { ServicesProvider, useServices } from '../contexts/ServicesContext';
import { SettingsProvider } from '../contexts/SettingsContext';
import { WizardProvider, isDraftDirty, useWizard } from '../contexts/WizardContext';
import { OffboardProvider, isOffboardDirty, useOffboard } from '../contexts/OffboardContext';
import { useAppRoles } from '../hooks/useReferenceData';
import { AppToasterProvider } from './Shared/AppToaster';
import { AppErrorBoundary } from './Shared/AppErrorBoundary';
import { PreflightBar } from './Preflight/PreflightBar';
import { JobsList } from './Jobs/JobsList';
import { OnboardingWizard } from './Onboarding/OnboardingWizard';
import { OffboardingWizard } from './Offboarding/OffboardingWizard';

// Tab content is lazy-loaded so the initial bundle stays small; rarely-used
// tabs (Tasks, Templates, Settings, Bulk, Transfer) only load when first opened.
const TransferPanel = React.lazy(() =>
  import(/* webpackChunkName: 'upc-transfer' */ './Transfer/TransferPanel').then((m) => ({
    default: m.TransferPanel
  }))
);
const BulkImport = React.lazy(() =>
  import(/* webpackChunkName: 'upc-bulk' */ './Bulk/BulkImport').then((m) => ({ default: m.BulkImport }))
);
const BulkOffboard = React.lazy(() =>
  import(/* webpackChunkName: 'upc-bulk-offboard' */ './Bulk/BulkOffboard').then((m) => ({
    default: m.BulkOffboard
  }))
);
const TasksList = React.lazy(() =>
  import(/* webpackChunkName: 'upc-tasks' */ './Tasks/TasksList').then((m) => ({ default: m.TasksList }))
);
const TemplatesList = React.lazy(() =>
  import(/* webpackChunkName: 'upc-templates' */ './Templates/TemplatesList').then((m) => ({ default: m.TemplatesList }))
);
const SettingsPanel = React.lazy(() =>
  import(/* webpackChunkName: 'upc-settings' */ './Settings/SettingsPanel').then((m) => ({ default: m.SettingsPanel }))
);
const RolesPanel = React.lazy(() =>
  import(/* webpackChunkName: 'upc-roles' */ './Roles/RolesPanel').then((m) => ({ default: m.RolesPanel }))
);

const useStyles = makeStyles({
  root: {
    // The web part must not inherit the page's typography: pin the Fluent 2
    // base ramp (Segoe UI, Body1) on the root.
    fontFamily: tokens.fontFamilyBase,
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    color: tokens.colorNeutralForeground1,
    backgroundColor: tokens.colorNeutralBackground1,
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalL,
    // Guest on the page: no fixed positioning, respects section width.
    width: '100%',
    // Descendants respond to the web part's width, not the viewport's —
    // section columns can be narrow on wide screens.
    containerType: 'inline-size'
  },
  header: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalXS
  },
  headerCaption: {
    color: tokens.colorNeutralForeground3
  },
  tabs: {
    borderBottomWidth: tokens.strokeWidthThin,
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2
  },
  // Reserves space so the transition indicator appearing/disappearing never
  // shifts layout.
  progressSlot: {
    minHeight: '2px'
  },
  draftBadge: {
    marginLeft: tokens.spacingHorizontalXS
  }
});

type ShellTab =
  | 'dashboard'
  | 'newUser'
  | 'offboard'
  | 'transfer'
  | 'bulk'
  | 'bulkOffboard'
  | 'tasks'
  | 'templates'
  | 'settings'
  | 'roles';

const DraftBadge: React.FC<{ className: string }> = ({ className }) => (
  <Badge className={className} appearance="tint" color="brand" size="small">
    {strings.TabDraftBadge}
  </Badge>
);

const Shell: React.FC = () => {
  const styles = useStyles();
  const [tab, setTab] = React.useState<ShellTab>('dashboard');
  // Tab switches are marked as transitions (React 18): the clicked tab
  // highlights immediately, while mounting the destination's content (a
  // wizard, a DataGrid) is deprioritized instead of blocking that feedback.
  const [isTabPending, startTabTransition] = React.useTransition();
  const { state: wizardState } = useWizard();
  const { state: offboardState } = useOffboard();
  const roles = useAppRoles();
  const services = useServices();
  const onRenderError = React.useCallback(
    (error: Error, info: React.ErrorInfo) =>
      services.telemetry.trackError(error, { componentStack: info.componentStack, scope: 'render' }),
    [services]
  );

  const onboardDraft: boolean = isDraftDirty(wizardState);
  const offboardDraft: boolean = isOffboardDirty(offboardState);
  const canManageTemplates: boolean = roles.data?.permissions.has('manageTemplates') ?? false;
  const canManageTasks: boolean = roles.data?.permissions.has('manageTasks') ?? false;
  const canManageSettings: boolean = roles.data?.permissions.has('manageSettings') ?? false;
  const switchTab = (next: ShellTab): void => {
    startTabTransition(() => setTab(next));
  };
  const goToDashboard = (): void => switchTab('dashboard');

  return (
    <div className={styles.root}>
      <PreflightBar />
      <header className={styles.header}>
        <Title3 as="h2" block>
          {strings.AppTitle}
        </Title3>
        <Caption1 block className={styles.headerCaption}>
          {strings.AppDescription}
        </Caption1>
      </header>
      <div className={styles.tabs}>
        <TabList
          selectedValue={tab}
          onTabSelect={(_, data) => switchTab(data.value as ShellTab)}
          aria-label={strings.AppTitle}
        >
          <Tab value="dashboard">{strings.TabDashboard}</Tab>
          <Tab value="newUser">
            {strings.TabNewUser}
            {onboardDraft ? <DraftBadge className={styles.draftBadge} /> : undefined}
          </Tab>
          <Tab value="offboard">
            {strings.TabOffboard}
            {offboardDraft ? <DraftBadge className={styles.draftBadge} /> : undefined}
          </Tab>
          <Tab value="transfer">{strings.TabTransfer}</Tab>
          <Tab value="bulk">{strings.TabBulk}</Tab>
          <Tab value="bulkOffboard">{strings.TabBulkOffboard}</Tab>
          {canManageTasks ? <Tab value="tasks">{strings.TabTasks}</Tab> : undefined}
          {canManageTemplates ? <Tab value="templates">{strings.TabTemplates}</Tab> : undefined}
          {canManageSettings ? <Tab value="settings">{strings.TabSettings}</Tab> : undefined}
          {canManageSettings ? <Tab value="roles">{strings.TabRoles}</Tab> : undefined}
        </TabList>
      </div>
      <div className={styles.progressSlot}>
        {isTabPending ? <ProgressBar aria-label={strings.LoadingLabel} /> : undefined}
      </div>
      <AppErrorBoundary onError={onRenderError}>
        {tab === 'dashboard' ? <JobsList onCreateNew={() => switchTab('newUser')} /> : undefined}
        {tab === 'newUser' ? <OnboardingWizard onSubmitted={goToDashboard} /> : undefined}
        {tab === 'offboard' ? <OffboardingWizard onSubmitted={goToDashboard} /> : undefined}
        {tab === 'transfer' ? (
          <React.Suspense fallback={<Spinner aria-label={strings.LoadingLabel} />}>
            <TransferPanel onSubmitted={goToDashboard} />
          </React.Suspense>
        ) : undefined}
        {tab === 'bulk' ? (
          <React.Suspense fallback={<Spinner aria-label={strings.LoadingLabel} />}>
            <BulkImport onSubmitted={goToDashboard} />
          </React.Suspense>
        ) : undefined}
        {tab === 'bulkOffboard' ? (
          <React.Suspense fallback={<Spinner aria-label={strings.LoadingLabel} />}>
            <BulkOffboard onSubmitted={goToDashboard} />
          </React.Suspense>
        ) : undefined}
        {tab === 'tasks' && canManageTasks ? (
          <React.Suspense fallback={<Spinner aria-label={strings.LoadingLabel} />}>
            <TasksList />
          </React.Suspense>
        ) : undefined}
        {tab === 'templates' && canManageTemplates ? (
          <React.Suspense fallback={<Spinner aria-label={strings.LoadingLabel} />}>
            <TemplatesList />
          </React.Suspense>
        ) : undefined}
        {tab === 'settings' && canManageSettings ? (
          <React.Suspense fallback={<Spinner aria-label={strings.LoadingLabel} />}>
            <SettingsPanel />
          </React.Suspense>
        ) : undefined}
        {tab === 'roles' && canManageSettings ? (
          <React.Suspense fallback={<Spinner aria-label={strings.LoadingLabel} />}>
            <RolesPanel />
          </React.Suspense>
        ) : undefined}
      </AppErrorBoundary>
    </div>
  );
};

export interface IAppProps {
  services: IServices;
  /** Bridged host theme (site theme variant / Teams). Falls back to web light. */
  theme?: Theme;
  dir?: 'ltr' | 'rtl';
  /** SPFx web part instance id — namespaces generated DOM ids and classes. */
  instanceId?: string;
}

/**
 * Root component. The theme is bridged from the host by the web part class
 * (section theme variants, Teams dark/high contrast); wizard providers live
 * here so in-progress drafts survive switching tabs.
 *
 * IdPrefixProvider MUST wrap FluentProvider: SharePoint Online's own chrome
 * runs Fluent v9, and without a unique prefix both bundles mint the same
 * `fui-FluentProviderX` class — portal surfaces (Dialog, Drawer, Toaster,
 * listboxes) then resolve their CSS variables against SharePoint's provider
 * and render unstyled (SharePoint/sp-dev-docs#9847, microsoft/fluentui#32253).
 */
export const App: React.FC<IAppProps> = ({ services, theme, dir, instanceId }) => {
  const [queryClient] = React.useState(
    () =>
      new QueryClient({
        defaultOptions: {
          queries: { retry: 1, refetchOnWindowFocus: false }
        }
      })
  );
  return (
    <IdPrefixProvider value={`upc-${instanceId ?? 'app'}-`}>
      <FluentProvider theme={theme ?? webLightTheme} dir={dir ?? 'ltr'}>
        <QueryClientProvider client={queryClient}>
          <ServicesProvider services={services}>
            <SettingsProvider>
              <AppToasterProvider>
                <WizardProvider>
                  <OffboardProvider>
                    <Shell />
                  </OffboardProvider>
                </WizardProvider>
              </AppToasterProvider>
            </SettingsProvider>
          </ServicesProvider>
        </QueryClientProvider>
      </FluentProvider>
    </IdPrefixProvider>
  );
};
