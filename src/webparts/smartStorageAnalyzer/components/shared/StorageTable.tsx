import * as React from 'react';
import { makeStyles, tokens } from '@fluentui/react-components';
import { ChevronDown16Regular, ChevronRight16Regular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  table: {
    width: '100%',
    borderCollapse: 'collapse',
    fontSize: tokens.fontSizeBase200,
  },
  th: {
    textAlign: 'left',
    padding: '6px 8px',
    borderBottom: `2px solid ${tokens.colorNeutralStroke1}`,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    whiteSpace: 'nowrap',
    position: 'sticky',
    top: 0,
    background: tokens.colorNeutralBackground1,
    zIndex: 1,
    cursor: 'pointer',
    userSelect: 'none',
  },
  thStatic: {
    cursor: 'default',
  },
  td: {
    padding: '5px 8px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    verticalAlign: 'middle',
  },
});

export interface StorageTableColumn<T> {
  key: string;
  header: string;
  render: (row: T) => React.ReactNode;
  sortValue?: (row: T) => number | string;
  align?: 'left' | 'right' | 'center';
}

export interface StorageTableProps<T> {
  rows: T[];
  columns: StorageTableColumn<T>[];
  getRowKey: (row: T) => string;
  defaultSortKey?: string;
  defaultSortDir?: 'asc' | 'desc';
}

// Generic hand-rolled sortable table (sticky header, click-to-sort columns)
// reused by the Library Overview list and the Storage Report / file-list
// results — deliberately not Fluent's DataGrid, to stay consistent with the
// sibling project's plain-<table> approach.
export function StorageTable<T>({ rows, columns, getRowKey, defaultSortKey, defaultSortDir = 'desc' }: StorageTableProps<T>): React.ReactElement {
  const styles = useStyles();
  const [sortKey, setSortKey] = React.useState<string | undefined>(defaultSortKey);
  const [sortDir, setSortDir] = React.useState<'asc' | 'desc'>(defaultSortDir);

  const sorted = React.useMemo(() => {
    const col = columns.find((c) => c.key === sortKey);
    if (!col?.sortValue) return rows;
    const copy = [...rows];
    copy.sort((a, b) => {
      const av = col.sortValue!(a);
      const bv = col.sortValue!(b);
      const cmp = typeof av === 'number' && typeof bv === 'number' ? av - bv : String(av).localeCompare(String(bv));
      return sortDir === 'asc' ? cmp : -cmp;
    });
    return copy;
  }, [rows, columns, sortKey, sortDir]);

  const handleHeaderClick = (col: StorageTableColumn<T>): void => {
    if (!col.sortValue) return;
    if (sortKey === col.key) {
      setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortKey(col.key);
      setSortDir('desc');
    }
  };

  return (
    <table className={styles.table}>
      <thead>
        <tr>
          {columns.map((col) => (
            <th
              key={col.key}
              className={`${styles.th}${col.sortValue ? '' : ` ${styles.thStatic}`}`}
              style={{ textAlign: col.align ?? 'left' }}
              onClick={() => handleHeaderClick(col)}
              role={col.sortValue ? 'button' : undefined}
              tabIndex={col.sortValue ? 0 : undefined}
              onKeyDown={col.sortValue ? (e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                  e.preventDefault();
                  handleHeaderClick(col);
                }
              } : undefined}
              aria-sort={col.sortValue ? (sortKey === col.key ? (sortDir === 'asc' ? 'ascending' : 'descending') : 'none') : undefined}
            >
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: '2px' }}>
                {col.header}
                {sortKey === col.key && (sortDir === 'asc' ? <ChevronRight16Regular style={{ transform: 'rotate(-90deg)' }} /> : <ChevronDown16Regular />)}
              </span>
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {sorted.map((row) => (
          <tr key={getRowKey(row)}>
            {columns.map((col) => (
              <td key={col.key} className={styles.td} style={{ textAlign: col.align ?? 'left' }}>
                {col.render(row)}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
