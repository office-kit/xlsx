// xlsx streaming entry point — read-only iter + write-only append.
// Format-agnostic byte I/O lives at `@office-kit/xlsx/io` and `@office-kit/xlsx/node`;
// error types at `@office-kit/xlsx/utils`.

export {
  loadWorkbookStream,
  type IterRowsOptions,
  type LoadWorkbookStreamOptions,
  type ReadOnlyCell,
  type ReadOnlyWorkbook,
  type ReadOnlyWorksheet,
} from './read-only';

export {
  createWriteOnlyWorkbook,
  type WriteOnlyOptions,
  type WriteOnlyRowItem,
  type WriteOnlyStyle,
  type WriteOnlyWorkbook,
  type WriteOnlyWorksheet,
} from './write-only';
