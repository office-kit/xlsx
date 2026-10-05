<script lang="ts">
  // Renders the dialog the controller asks for. Each dialog is keyed by its
  // state object so reopening the same kind starts from fresh state.
  import type { Component } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import type { DialogKind } from '../core/controller.svelte.ts';
  import AlertDialog from './AlertDialog.svelte';
  import AltTextDialog from './AltTextDialog.svelte';
  import ConditionalFormattingDialog from './ConditionalFormattingDialog.svelte';
  import CreatePivotTableDialog from './CreatePivotTableDialog.svelte';
  import CreateSparklinesDialog from './CreateSparklinesDialog.svelte';
  import CreateTableDialog from './CreateTableDialog.svelte';
  import DataValidationDialog from './DataValidationDialog.svelte';
  import DefineNameDialog from './DefineNameDialog.svelte';
  import FindReplaceDialog from './FindReplaceDialog.svelte';
  import FormatCellsDialog from './FormatCellsDialog.svelte';
  import GoToDialog from './GoToDialog.svelte';
  import GoToSpecialDialog from './GoToSpecialDialog.svelte';
  import HyperlinkDialog from './HyperlinkDialog.svelte';
  import InsertChartDialog from './InsertChartDialog.svelte';
  import InsertDeleteCellsDialog from './InsertDeleteCellsDialog.svelte';
  import InsertFunctionDialog from './InsertFunctionDialog.svelte';
  import MoveChartDialog from './MoveChartDialog.svelte';
  import MoveCopySheetDialog from './MoveCopySheetDialog.svelte';
  import NameManagerDialog from './NameManagerDialog.svelte';
  import NoteDialog from './NoteDialog.svelte';
  import PageSetupDialog from './PageSetupDialog.svelte';
  import PasteSpecialDialog from './PasteSpecialDialog.svelte';
  import ProtectSheetDialog from './ProtectSheetDialog.svelte';
  import ProtectWorkbookDialog from './ProtectWorkbookDialog.svelte';
  import UnprotectDialog from './UnprotectDialog.svelte';
  import EvaluateFormulaDialog from './EvaluateFormulaDialog.svelte';
  import SymbolDialog from './SymbolDialog.svelte';
  import WorkbookStatisticsDialog from './WorkbookStatisticsDialog.svelte';
  import RemoveDuplicatesDialog from './RemoveDuplicatesDialog.svelte';
  import RenameSheetDialog from './RenameSheetDialog.svelte';
  import SelectDataDialog from './SelectDataDialog.svelte';
  import SeriesDialog from './SeriesDialog.svelte';
  import SizeDialog from './SizeDialog.svelte';
  import SortDialog from './SortDialog.svelte';
  import TextToColumnsDialog from './TextToColumnsDialog.svelte';
  import UnhideSheetDialog from './UnhideSheetDialog.svelte';
  import ZoomDialog from './ZoomDialog.svelte';
  import ResizeTableDialog from './ResizeTableDialog.svelte';
  import SubtotalDialog from './SubtotalDialog.svelte';
  import ConsolidateDialog from './ConsolidateDialog.svelte';
  import AdvancedFilterDialog from './AdvancedFilterDialog.svelte';
  import GoalSeekDialog from './GoalSeekDialog.svelte';
  import ScenarioManagerDialog from './ScenarioManagerDialog.svelte';
  import DataTableDialog from './DataTableDialog.svelte';

  type DialogProps = { props: Record<string, unknown> | undefined };

  // Dialogs that read their opening arguments (a sheet index, the tab to show,
  // or — for dialogs serving several kinds — `props.kind`).
  const WITH_PROPS: Partial<Record<DialogKind, Component<DialogProps>>> = {
    formatCells: FormatCellsDialog,
    find: FindReplaceDialog,
    replace: FindReplaceDialog,
    insertCells: InsertDeleteCellsDialog,
    deleteCells: InsertDeleteCellsDialog,
    columnWidth: SizeDialog,
    rowHeight: SizeDialog,
    standardWidth: SizeDialog,
    renameSheet: RenameSheetDialog,
    moveCopySheet: MoveCopySheetDialog,
    conditionalFormatting: ConditionalFormattingDialog,
    defineName: DefineNameDialog,
    pageSetup: PageSetupDialog,
    insertChart: InsertChartDialog,
    alert: AlertDialog,
    createSparklines: CreateSparklinesDialog,
    createPivotTable: CreatePivotTableDialog,
    resizeTable: ResizeTableDialog,
    unprotect: UnprotectDialog,
  };
  const PLAIN: Partial<Record<DialogKind, Component>> = {
    goto: GoToDialog,
    gotoSpecial: GoToSpecialDialog,
    pasteSpecial: PasteSpecialDialog,
    sort: SortDialog,
    dataValidation: DataValidationDialog,
    nameManager: NameManagerDialog,
    insertFunction: InsertFunctionDialog,
    hyperlink: HyperlinkDialog,
    note: NoteDialog,
    createTable: CreateTableDialog,
    chartSelectData: SelectDataDialog,
    moveChart: MoveChartDialog,
    altText: AltTextDialog,
    zoom: ZoomDialog,
    unhideSheet: UnhideSheetDialog,
    protectSheet: ProtectSheetDialog,
    protectWorkbook: ProtectWorkbookDialog,
    evaluateFormula: EvaluateFormulaDialog,
    symbol: SymbolDialog,
    workbookStatistics: WorkbookStatisticsDialog,
    removeDuplicates: RemoveDuplicatesDialog,
    textToColumns: TextToColumnsDialog,
    series: SeriesDialog,
    subtotal: SubtotalDialog,
    consolidate: ConsolidateDialog,
    advancedFilter: AdvancedFilterDialog,
    goalSeek: GoalSeekDialog,
    scenarioManager: ScenarioManagerDialog,
    dataTable: DataTableDialog,
  };

  const ctl = getEditor();
</script>

{#if ctl.dialog}
  {@const dialog = ctl.dialog}
  {@const WithProps = WITH_PROPS[dialog.kind]}
  {@const Plain = PLAIN[dialog.kind]}
  {#key dialog}
    {#if WithProps}
      <WithProps props={{ ...dialog.props, kind: dialog.kind }} />
    {:else if Plain}
      <Plain />
    {/if}
  {/key}
{/if}
