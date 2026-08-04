import * as React from 'react';
import { makeStyles, tokens, Button, Text } from '@fluentui/react-components';
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
  // When set, only this many rows render into the DOM at once, with
  // Previous/Next paging controls below the table. Without this, a report
  // with tens or hundreds of thousands of rows (a full-site Storage Report
  // scan) creates that many <tr> elements in one go — slow to paint, and
  // heavy enough on a real 190,000-row report to make the tab stutter just
  // scrolling it. Omit for tables that are never realistically that large
  // (e.g. one folder's immediate children) to keep today's unpaged behavior.
  pageSize?: number;
}

// Generic hand-rolled sortable table (sticky header, click-to-sort columns)
// reused by the Library Overview list and the Storage Report / file-list
// results — deliberately not Fluent's DataGrid, to stay consistent with the
// sibling project's plain-<table> approach.
export function StorageTable<T>({
  rows, columns, getRowKey, defaultSortKey, defaultSortDir = 'desc', pageSize,
}: StorageTableProps<T>): React.ReactElement {
  const styles = useStyles();
  const [sortKey, setSortKey] = React.useState<string | undefined>(defaultSortKey);
  const [sortDir, setSortDir] = React.useState<'asc' | 'desc'>(defaultSortDir);
  const [page, setPage] = React.useState(0);

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

  // Reset to page 1 whenever the data or sort changes — otherwise re-scanning,
  // re-sorting, or toggling a filter (which swaps `rows` for a shorter array)
  // could leave the user stranded on a page past the end of the new results.
  React.useEffect(() => { setPage(0); }, [rows, sortKey, sortDir]);

  const pageCount = pageSize ? Math.max(1, Math.ceil(sorted.length / pageSize)) : 1;
  const currentPage = Math.min(page, pageCount - 1);
  const visible = pageSize ? sorted.slice(currentPage * pageSize, currentPage * pageSize + pageSize) : sorted;

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
    <>
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
        {visible.map((row) => (
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
    {pageSize != null && sorted.length > pageSize && (
      <div style={{
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        flexWrap: 'wrap',
        gap: tokens.spacingHorizontalM,
        marginTop: tokens.spacingVerticalS,
      }}>
        <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
          Showing {(currentPage * pageSize + 1).toLocaleString()}–
          {Math.min((currentPage + 1) * pageSize, sorted.length).toLocaleString()} of {sorted.length.toLocaleString()}
        </Text>
        <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
          <Button size="small" appearance="secondary" disabled={currentPage === 0} onClick={() => setPage((p) => Math.max(0, p - 1))}>
            Previous
          </Button>
          <Text style={{ fontSize: tokens.fontSizeBase200, whiteSpace: 'nowrap' }}>
            Page {currentPage + 1} of {pageCount}
          </Text>
          <Button
            size="small"
            appearance="secondary"
            disabled={currentPage >= pageCount - 1}
            onClick={() => setPage((p) => Math.min(pageCount - 1, p + 1))}
          >
            Next
          </Button>
        </div>
      </div>
    )}
    </>
  );
}
