// Maps an OPC content type to the XSD that covers it.
//
// The schemas in `schemas/strict/`, `schemas/transitional/` and `schemas/opc/`
// are vendored from ECMA-376 (Part 1 Strict, Part 4 Transitional, Part 2 OPC). xmllint resolves
// each schema's relative `<xsd:import>` paths against the file's own
// directory, so cross-schema references (e.g. sml.xsd → dml-spreadsheetDrawing.xsd)
// work as long as we point xmllint at the leaf XSD listed below.
//
// Content types not present here are intentionally skipped during XSD
// validation — they're either binary parts (printer settings, OLE blobs)
// or formats our library does not generate yet.

import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const TRANSITIONAL_DIR = resolve(HERE, 'schemas/transitional');
const OPC_DIR = resolve(HERE, 'schemas/opc');

const SML = join(TRANSITIONAL_DIR, 'sml.xsd');
const DML_MAIN = join(TRANSITIONAL_DIR, 'dml-main.xsd');
const DML_CHART = join(TRANSITIONAL_DIR, 'dml-chart.xsd');
const DML_CHART_DRAWING = join(TRANSITIONAL_DIR, 'dml-chartDrawing.xsd');
const DML_DIAGRAM = join(TRANSITIONAL_DIR, 'dml-diagram.xsd');
const DML_SPREADSHEET_DRAWING = join(TRANSITIONAL_DIR, 'dml-spreadsheetDrawing.xsd');
const SHARED_EXTENDED = join(TRANSITIONAL_DIR, 'shared-documentPropertiesExtended.xsd');
const SHARED_CUSTOM = join(TRANSITIONAL_DIR, 'shared-documentPropertiesCustom.xsd');
const OPC_CONTENT_TYPES = join(OPC_DIR, 'opc-contentTypes.xsd');
const OPC_CORE_PROPERTIES = join(OPC_DIR, 'opc-coreProperties.xsd');
const OPC_RELATIONSHIPS = join(OPC_DIR, 'opc-relationships.xsd');

const SHEETML = 'application/vnd.openxmlformats-officedocument.spreadsheetml';
const ODOC = 'application/vnd.openxmlformats-officedocument';
const PKG = 'application/vnd.openxmlformats-package';

/** Maps OPC content type → absolute path of the XSD root file. */
const SCHEMA_BY_CONTENT_TYPE: Readonly<Record<string, string>> = {
  // SpreadsheetML core
  'application/vnd.ms-excel.sheet.macroEnabled.main+xml': SML,
  'application/vnd.ms-excel.template.macroEnabled.main+xml': SML,
  [`${SHEETML}.sheet.main+xml`]: SML,
  [`${SHEETML}.template.main+xml`]: SML,
  [`${SHEETML}.worksheet+xml`]: SML,
  [`${SHEETML}.chartsheet+xml`]: SML,
  [`${SHEETML}.dialogsheet+xml`]: SML,
  [`${SHEETML}.styles+xml`]: SML,
  [`${SHEETML}.sharedStrings+xml`]: SML,
  [`${SHEETML}.comments+xml`]: SML,
  [`${SHEETML}.table+xml`]: SML,
  [`${SHEETML}.tableSingleCells+xml`]: SML,
  [`${SHEETML}.queryTable+xml`]: SML,
  [`${SHEETML}.connections+xml`]: SML,
  [`${SHEETML}.externalLink+xml`]: SML,
  [`${SHEETML}.pivotTable+xml`]: SML,
  [`${SHEETML}.pivotCacheDefinition+xml`]: SML,
  [`${SHEETML}.pivotCacheRecords+xml`]: SML,
  [`${SHEETML}.calcChain+xml`]: SML,
  [`${SHEETML}.revisionHeaders+xml`]: SML,
  [`${SHEETML}.revisionLog+xml`]: SML,
  [`${SHEETML}.userNames+xml`]: SML,
  [`${SHEETML}.volatileDependencies+xml`]: SML,

  // DrawingML
  [`${ODOC}.theme+xml`]: DML_MAIN,
  [`${ODOC}.themeOverride+xml`]: DML_MAIN,
  [`${ODOC}.drawing+xml`]: DML_SPREADSHEET_DRAWING,
  [`${ODOC}.drawingml.chart+xml`]: DML_CHART,
  [`${ODOC}.drawingml.chartshapes+xml`]: DML_CHART_DRAWING,
  [`${ODOC}.drawingml.diagramData+xml`]: DML_DIAGRAM,
  [`${ODOC}.drawingml.diagramLayout+xml`]: DML_DIAGRAM,
  [`${ODOC}.drawingml.diagramStyle+xml`]: DML_DIAGRAM,
  [`${ODOC}.drawingml.diagramColors+xml`]: DML_DIAGRAM,

  // Document properties
  [`${ODOC}.extended-properties+xml`]: SHARED_EXTENDED,
  [`${ODOC}.custom-properties+xml`]: SHARED_CUSTOM,
  [`${PKG}.core-properties+xml`]: OPC_CORE_PROPERTIES,

  // Package-level
  [`${PKG}.relationships+xml`]: OPC_RELATIONSHIPS,
};

/** Path to the OPC content-types schema (used to validate `[Content_Types].xml`). */
export const CONTENT_TYPES_SCHEMA = OPC_CONTENT_TYPES;
/** Path to the OPC relationships schema (used to validate every `*.rels`). */
export const RELATIONSHIPS_SCHEMA = OPC_RELATIONSHIPS;

/** Returns the schema path for a content type, or undefined if unmapped. */
export function schemaFor(contentType: string, conformance: 'transitional' | 'strict' = 'transitional'): string | undefined {
  const schema = SCHEMA_BY_CONTENT_TYPE[contentType];
  return schema && conformance === 'strict' && dirname(schema) === TRANSITIONAL_DIR
    ? join(HERE, 'schemas/strict', basename(schema)) : schema;
}
