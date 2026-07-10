import * as React from 'react';
import {
  Button,
  DataGrid,
  DataGridBody,
  DataGridCell,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridRow,
  Dropdown,
  MessageBar,
  MessageBarBody,
  Option,
  SearchBox,
  Spinner,
  Subtitle2,
  Text,
  ToolbarButton,
  createTableColumn,
  makeStyles,
  shorthands,
  tokens,
  type TableColumnDefinition
} from '@fluentui/react-components';
import { ArrowClockwise16Regular, ArrowDownload16Regular, CheckmarkRegular } from '@fluentui/react-icons';
import { useQueryClient } from '@tanstack/react-query';
import * as strings from 'UpcStrings';
import { useAppRoles, useJobs } from '../../hooks/useReferenceData';
import { useServices } from '../../contexts/ServicesContext';
import { QK_JOBS } from '../../constants/queryKeys';
import { isOnboardingPayload } from '../../models';
import type { IProvisioningJob, JobType } from '../../models';
import { targetUserOf } from '../../services/engine/stepHelpers';
import { downloadCsv, toCsv } from '../../services/util/csv';
import { DataState } from '../Shared/DataState';
import { useAppToast } from '../Shared/AppToaster';
import { JobStatusBadge, jobTypeLabel } from '../Shared/StatusBadge';
import { JobDetailDrawer } from './JobDetailDrawer';
import type { KpiFilterKey } from './DashboardOverview';

// recharts (pulled in by DashboardOverview) is the single largest dependency
// this tab touches; lazy-load it behind Suspense like the other secondary
// tabs in App.tsx rather than shipping it in the eager/default bundle.
const DashboardOverview = React.lazy(() =>
  import(/* webpackChunkName: 'upc-charts' */ './DashboardOverview').then((m) => ({
    default: m.DashboardOverview
  }))
);

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalM
  },
  titleRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between'
  },
  filterRow: {
    display: 'flex',
    flexWrap: 'wrap',
    columnGap: tokens.spacingHorizontalS,
    rowGap: tokens.spacingVerticalS,
    alignItems: 'center'
  },
  search: {
    minWidth: '220px',
    flexGrow: 1,
    maxWidth: '360px'
  },
  searchStale: {
    marginLeft: tokens.spacingHorizontalXS
  },
  // Dimmed while a deferred search re-filter is still catching up with the
  // input — keeps typing itself perfectly responsive (React 18 useDeferredValue).
  resultsStale: {
    opacity: 0.6,
    transitionProperty: 'opacity',
    transitionDuration: tokens.durationSlower,
    transitionTimingFunction: tokens.curveEasyEase
  },
  tableWrap: {
    overflowX: 'auto',
    ...shorthands.border(tokens.strokeWidthThin, 'solid', tokens.colorNeutralStroke2),
    borderRadius: tokens.borderRadiusMedium
  },
  row: {
    cursor: 'pointer',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover
    }
  },
  targetButton: {
    paddingLeft: 0,
    paddingRight: 0,
    fontWeight: tokens.fontWeightSemibold
  },
  noMatches: {
    color: tokens.colorNeutralForeground3,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalM
  }
});

const WEEK_MS: number = 7 * 24 * 60 * 60 * 1000;

const JOB_TYPES: readonly JobType[] = ['Onboard', 'Offboard', 'Transfer', 'Clone', 'Bulk'];

type KpiKey = 'pending' | 'running' | 'failed7' | 'completedToday';

function targetOf(job: IProvisioningJob): string {
  return targetUserOf(job.payload) || job.jobId.slice(0, 8);
}

/**
 * True when the template this job was created from restricts approval to a
 * specific Entra group. The dashboard grid's inline Approve is a one-click
 * shortcut with no room for a membership check against N rows at once — a
 * restricted job is approved from the job drawer instead, where the check
 * actually runs (see JobDetailDrawer's showApprove).
 */
function hasApproverRestriction(job: IProvisioningJob): boolean {
  return isOnboardingPayload(job.payload) && !!job.payload.approverGroupId;
}

function createdLabel(job: IProvisioningJob): string {
  if (!job.createdUtc) {
    return '';
  }
  const date: Date = new Date(job.createdUtc);
  return isNaN(date.getTime()) ? '' : date.toLocaleString();
}

function isWithinWeek(job: IProvisioningJob, now: number): boolean {
  return !!job.createdUtc && now - Date.parse(job.createdUtc) <= WEEK_MS;
}

const DAY_MS: number = 24 * 60 * 60 * 1000;

function matchesKpi(job: IProvisioningJob, kpi: KpiKey, now: number): boolean {
  switch (kpi) {
    case 'pending':
      return job.status === 'PendingApproval';
    case 'running':
      return job.status === 'Running';
    case 'failed7':
      return (
        (job.status === 'Failed' || job.status === 'PartiallyFailed') && isWithinWeek(job, now)
      );
    case 'completedToday':
      return (
        job.status === 'Completed' &&
        !!job.createdUtc &&
        now - Date.parse(job.createdUtc) <= DAY_MS
      );
    default:
      return true;
  }
}

export interface IJobsListProps {
  /** Empty-state call to action: jump to the New user tab. */
  onCreateNew: () => void;
}

/**
 * Dashboard: KPI counts that double as filters, search + type filter, and a
 * sortable person-first job grid. Opening a job exposes Run/Resume per its
 * persisted state.
 */
export const JobsList: React.FC<IJobsListProps> = ({ onCreateNew }) => {
  const styles = useStyles();
  const jobs = useJobs();
  const roles = useAppRoles();
  const services = useServices();
  const queryClient = useQueryClient();
  const toast = useAppToast();
  const [selectedItemId, setSelectedItemId] = React.useState<number | null>(null);
  const [approvingItemId, setApprovingItemId] = React.useState<number | null>(null);
  const [kpi, setKpi] = React.useState<KpiKey | null>(null);
  const [search, setSearch] = React.useState<string>('');
  const [typeFilter, setTypeFilter] = React.useState<JobType | 'all'>('all');

  const canApprove: boolean = roles.data?.permissions.has('approveJobs') ?? false;

  const approveInline = React.useCallback(
    async (itemId: number): Promise<void> => {
      setApprovingItemId(itemId);
      try {
        await services.data.approveJob(itemId);
        toast(strings.JobApprovedToast);
      } catch {
        toast(strings.JobActionFailed, 'error');
      } finally {
        setApprovingItemId(null);
        void queryClient.invalidateQueries(QK_JOBS);
      }
    },
    [services.data, queryClient, toast]
  );

  const items: IProvisioningJob[] = jobs.data?.items ?? [];
  const truncated: boolean = jobs.data?.truncated ?? false;
  const selected: IProvisioningJob | null =
    items.filter((j) => j.itemId === selectedItemId)[0] ?? null;

  const now: number = Date.now();

  // Typing stays instantly responsive; the (potentially large) filtered grid
  // is allowed to lag a frame behind under load (React 18 useDeferredValue).
  const deferredSearch: string = React.useDeferredValue(search);
  const isSearchStale: boolean = search !== deferredSearch;
  const needle: string = deferredSearch.trim().toLowerCase();
  const filtered: IProvisioningJob[] = items
    .filter((j) => (kpi ? matchesKpi(j, kpi, now) : true))
    .filter((j) => (typeFilter === 'all' ? true : j.jobType === typeFilter))
    .filter(
      (j) =>
        !needle ||
        targetOf(j).toLowerCase().indexOf(needle) !== -1 ||
        (j.requestedBy ?? '').toLowerCase().indexOf(needle) !== -1
    );

  const columns: TableColumnDefinition<IProvisioningJob>[] = React.useMemo(
    () => [
      createTableColumn<IProvisioningJob>({
        columnId: 'target',
        compare: (a, b) => targetOf(a).localeCompare(targetOf(b)),
        renderHeaderCell: () => strings.JobColumnTarget,
        renderCell: (job) => (
          <Button
            appearance="transparent"
            className={styles.targetButton}
            onClick={() => setSelectedItemId(job.itemId)}
          >
            {targetOf(job)}
          </Button>
        )
      }),
      createTableColumn<IProvisioningJob>({
        columnId: 'type',
        compare: (a, b) => a.jobType.localeCompare(b.jobType),
        renderHeaderCell: () => strings.JobColumnType,
        renderCell: (job) => jobTypeLabel(job.jobType)
      }),
      createTableColumn<IProvisioningJob>({
        columnId: 'status',
        compare: (a, b) => a.status.localeCompare(b.status),
        renderHeaderCell: () => strings.JobColumnStatus,
        renderCell: (job) => <JobStatusBadge status={job.status} />
      }),
      createTableColumn<IProvisioningJob>({
        columnId: 'requestedBy',
        compare: (a, b) => (a.requestedBy ?? '').localeCompare(b.requestedBy ?? ''),
        renderHeaderCell: () => strings.JobColumnRequestedBy,
        renderCell: (job) => job.requestedBy ?? ''
      }),
      createTableColumn<IProvisioningJob>({
        columnId: 'created',
        compare: (a, b) => Date.parse(a.createdUtc ?? '') - Date.parse(b.createdUtc ?? ''),
        renderHeaderCell: () => strings.JobColumnCreated,
        renderCell: createdLabel
      }),
      // Inline approval: only rendered for operators holding approveJobs, and
      // only on rows still waiting for a decision.
      ...(canApprove
        ? [
            createTableColumn<IProvisioningJob>({
              columnId: 'actions',
              compare: () => 0,
              renderHeaderCell: () => strings.JobColumnActions,
              renderCell: (job) =>
                job.status === 'PendingApproval' && !hasApproverRestriction(job) ? (
                  <Button
                    size="small"
                    icon={<CheckmarkRegular />}
                    disabled={approvingItemId !== null}
                    onClick={(ev) => {
                      // The row itself opens the drawer; approving must not.
                      ev.stopPropagation();
                      void approveInline(job.itemId);
                    }}
                  >
                    {strings.ApproveLabel}
                  </Button>
                ) : null
            })
          ]
        : [])
    ],
    [styles.targetButton, canApprove, approvingItemId, approveInline]
  );

  const jobTypes: readonly JobType[] = JOB_TYPES;

  const exportCsv = (): void => {
    const csv: string = toCsv(
      [
        strings.JobColumnTarget,
        strings.JobColumnType,
        strings.JobColumnStatus,
        strings.JobColumnRequestedBy,
        strings.ApprovedByLabel,
        strings.JobColumnCreated,
        strings.CorrelationIdLabel
      ],
      filtered.map((job) => [
        targetOf(job),
        jobTypeLabel(job.jobType),
        job.status,
        job.requestedBy ?? '',
        job.approvedBy ?? '',
        createdLabel(job),
        job.correlationId
      ])
    );
    downloadCsv(`upc-jobs-${new Date().toISOString().slice(0, 10)}.csv`, csv);
  };

  return (
    <div className={styles.root}>
      <div className={styles.titleRow}>
        <Subtitle2 as="h3" block>
          {strings.JobsTitle}
        </Subtitle2>
        <ToolbarButton
          icon={<ArrowClockwise16Regular />}
          onClick={() => {
            void jobs.refetch();
          }}
          aria-label={strings.RefreshLabel}
        >
          {strings.RefreshLabel}
        </ToolbarButton>
        <ToolbarButton
          icon={<ArrowDownload16Regular />}
          disabled={filtered.length === 0}
          onClick={exportCsv}
          aria-label={strings.ExportCsvLabel}
        >
          {strings.ExportCsvLabel}
        </ToolbarButton>
      </div>
      <DataState
        isLoading={jobs.isLoading}
        error={jobs.error}
        isEmpty={!jobs.isLoading && items.length === 0}
        emptyTitle={strings.JobsEmptyTitle}
        emptyBody={strings.JobsEmptyBody}
        emptyAction={
          <Button appearance="secondary" onClick={onCreateNew}>
            {strings.JobsEmptyCta}
          </Button>
        }
        onRetry={() => {
          void jobs.refetch();
        }}
      >
        {items.length > 0 ? (
          <React.Suspense fallback={<Spinner aria-label={strings.LoadingLabel} />}>
            <DashboardOverview
              jobs={items}
              activeKpi={kpi as KpiFilterKey}
              onKpiClick={(key) => setKpi(key as KpiKey)}
            />
          </React.Suspense>
        ) : undefined}
        {truncated ? (
          <MessageBar intent="warning">
            <MessageBarBody>{strings.JobsTruncatedWarning}</MessageBarBody>
          </MessageBar>
        ) : undefined}
        <div className={styles.filterRow}>
          <SearchBox
            className={styles.search}
            placeholder={strings.SearchJobsPlaceholder}
            aria-label={strings.SearchJobsPlaceholder}
            value={search}
            onChange={(_, data) => setSearch(data.value)}
          />
          {isSearchStale ? (
            <Spinner
              size="tiny"
              className={styles.searchStale}
              aria-label={strings.LoadingLabel}
            />
          ) : undefined}
          <Dropdown
            value={typeFilter === 'all' ? strings.FilterTypeAll : jobTypeLabel(typeFilter)}
            selectedOptions={[typeFilter]}
            onOptionSelect={(_, data) =>
              setTypeFilter((data.optionValue as JobType | 'all') ?? 'all')
            }
            aria-label={strings.JobColumnType}
          >
            <Option value="all" text={strings.FilterTypeAll}>
              {strings.FilterTypeAll}
            </Option>
            {jobTypes.map((t) => (
              <Option key={t} value={t} text={jobTypeLabel(t)}>
                {jobTypeLabel(t)}
              </Option>
            ))}
          </Dropdown>
        </div>
        <div
          className={isSearchStale ? styles.resultsStale : undefined}
          aria-busy={isSearchStale}
        >
        {filtered.length === 0 ? (
          <Text className={styles.noMatches}>{strings.JobsNoMatches}</Text>
        ) : (
          <div className={styles.tableWrap}>
            <DataGrid
              items={filtered}
              columns={columns}
              sortable
              defaultSortState={{ sortColumn: 'created', sortDirection: 'descending' }}
              getRowId={(job: IProvisioningJob) => job.itemId}
              size="small"
              aria-label={strings.JobsTitle}
            >
              <DataGridHeader>
                <DataGridRow>
                  {({ renderHeaderCell }) => (
                    <DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>
                  )}
                </DataGridRow>
              </DataGridHeader>
              <DataGridBody<IProvisioningJob>>
                {({ item, rowId }) => (
                  <DataGridRow<IProvisioningJob>
                    key={rowId}
                    className={styles.row}
                    onClick={() => setSelectedItemId(item.itemId)}
                  >
                    {({ renderCell }) => <DataGridCell>{renderCell(item)}</DataGridCell>}
                  </DataGridRow>
                )}
              </DataGridBody>
            </DataGrid>
          </div>
        )}
        </div>
      </DataState>
      {selected ? (
        <JobDetailDrawer job={selected} onClose={() => setSelectedItemId(null)} />
      ) : undefined}
    </div>
  );
};
