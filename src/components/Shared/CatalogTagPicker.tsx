import * as React from 'react';
import {
  Field,
  Tag,
  TagPicker,
  TagPickerControl,
  TagPickerGroup,
  TagPickerInput,
  TagPickerList,
  TagPickerOption,
  Text,
  makeStyles,
  tokens,
  type TagPickerProps
} from '@fluentui/react-components';

const useStyles = makeStyles({
  empty: {
    padding: tokens.spacingVerticalS,
    color: tokens.colorNeutralForeground3
  },
  hint: {
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalSNudge}`,
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200
  }
});

export interface ICatalogTagPickerItem {
  id: string;
  title: string;
  /** Shown as secondary text under the option and, truncated, on its chip. */
  description?: string;
}

export interface ICatalogTagPickerProps {
  label: string;
  hint?: string;
  items: ICatalogTagPickerItem[];
  selectedIds: string[];
  onChange: (ids: string[]) => void;
  /** Shown instead of the picker when the catalog itself has no entries. */
  emptyMessage: string;
  /** Shown in the option list when a search matches nothing. */
  noMatchesMessage: string;
}

/**
 * Searchable, multi-select picker over an already-loaded catalog (Teams,
 * SharePoint sites, applications) — filters client-side since these lists
 * are small, curated SharePoint lists, not a live directory search. Same
 * chip UX as GroupTagPicker so every Access step section behaves the same way.
 */
export const CatalogTagPicker: React.FC<ICatalogTagPickerProps> = ({
  label,
  hint,
  items,
  selectedIds,
  onChange,
  emptyMessage,
  noMatchesMessage
}) => {
  const styles = useStyles();
  const [query, setQuery] = React.useState<string>('');
  const byId: Map<string, ICatalogTagPickerItem> = React.useMemo(
    () => new Map(items.map((item) => [item.id, item])),
    [items]
  );

  if (items.length === 0) {
    return (
      <Field label={label} hint={hint}>
        <Text className={styles.empty}>{emptyMessage}</Text>
      </Field>
    );
  }

  const needle: string = query.trim().toLowerCase();
  const options: ICatalogTagPickerItem[] = items.filter(
    (item) =>
      selectedIds.indexOf(item.id) === -1 &&
      (!needle || item.title.toLowerCase().indexOf(needle) !== -1)
  );

  const onOptionSelect: TagPickerProps['onOptionSelect'] = (_, data) => {
    onChange(data.selectedOptions);
    setQuery('');
  };

  return (
    <Field label={label} hint={hint}>
      <TagPicker selectedOptions={selectedIds} onOptionSelect={onOptionSelect}>
        <TagPickerControl>
          <TagPickerGroup aria-label={label}>
            {selectedIds.map((id) => (
              <Tag key={id} shape="rounded" value={id}>
                {byId.get(id)?.title ?? id}
              </Tag>
            ))}
          </TagPickerGroup>
          <TagPickerInput
            aria-label={label}
            value={query}
            onChange={(ev) => setQuery(ev.target.value)}
          />
        </TagPickerControl>
        <TagPickerList>
          {options.map((item) => (
            <TagPickerOption key={item.id} value={item.id} text={item.title} secondaryContent={item.description}>
              {item.title}
            </TagPickerOption>
          ))}
          {options.length === 0 ? <div className={styles.hint}>{noMatchesMessage}</div> : undefined}
        </TagPickerList>
      </TagPicker>
    </Field>
  );
};
