import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { packageFor, SML } from './corpus.js';
import { required } from './required.js';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
/** Hand-authored independent feature input; production serializers are not used. */
export function excelFeaturePackage(): Uint8Array {
  const parts = unzipSync(packageFor({ id: 'excel-table-chart-image-print', clause: 'ECMA-376-1', rowsXml: '<row r="1"><c r="A1" t="inlineStr"><is><t>audit</t></is></c></row><row r="2"><c r="A2"><v>1</v></c><c r="C2" t="inlineStr"><is><t>First</t></is></c></row><row r="3"><c r="A3"><v>2</v></c><c r="C3" t="inlineStr"><is><t>Second</t></is></c></row>' }));
  const sheet = 'xl/worksheets/sheet1.xml';
  parts[sheet] = strToU8(strFromU8(required(parts[sheet])).replace('<worksheet ', `<worksheet xmlns:r="${R}" `).replace('A1:C2', 'A1:C3').replace('</sheetData>', '</sheetData><printOptions horizontalCentered="1"/><pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.2" footer="0.2"/><pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/><drawing r:id="drawing"/><tableParts count="1"><tablePart r:id="table"/></tableParts>'));
  const xml: Record<string, string> = {
    'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${REL}"><Relationship Id="table" Type="${R}/table" Target="../tables/table1.xml"/><Relationship Id="drawing" Type="${R}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
    'xl/tables/table1.xml': `<table xmlns="${SML}" id="1" name="AuditTable" displayName="AuditTable" ref="A1:A3" totalsRowShown="0"><autoFilter ref="A1:A3"/><tableColumns count="1"><tableColumn id="1" name="audit"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`,
    'xl/charts/chart1.xml': `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:plotArea><c:layout/><c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strRef><c:f>Audit!$C$2:$C$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>First</c:v></c:pt><c:pt idx="1"><c:v>Second</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Audit!$A$2:$A$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:firstSliceAng val="0"/></c:pieChart></c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`,
    'xl/drawings/_rels/drawing1.xml.rels': `<Relationships xmlns="${REL}"><Relationship Id="chart" Type="${R}/chart" Target="../charts/chart1.xml"/><Relationship Id="image" Type="${R}/image" Target="../media/image1.png"/></Relationships>`,
  };
  const marker = (tag: string, col: number, row: number) => `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;
  xml['xl/drawings/drawing1.xml'] = `<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${R}"><xdr:twoCellAnchor>${marker('from', 4, 0)}${marker('to', 11, 11)}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Audit Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="chart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor><xdr:twoCellAnchor>${marker('from', 4, 14)}${marker('to', 6, 17)}<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Audit Image"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="image"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
  const types: Record<string, string> = {
    'xl/tables/table1.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
    'xl/drawings/drawing1.xml': 'application/vnd.openxmlformats-officedocument.drawing+xml',
    'xl/charts/chart1.xml': 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
    'xl/media/image1.png': 'image/png',
  };
  parts['[Content_Types].xml'] = strToU8(strFromU8(required(parts['[Content_Types].xml'])).replace('</Types>', Object.entries(types).map(([path, type]) => `<Override PartName="/${path}" ContentType="${type}"/>`).join('') + '</Types>'));
  for (const [path, content] of Object.entries(xml)) parts[path] = strToU8(content);
  parts['xl/media/image1.png'] = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAEAAAAAgCAIAAAAt/+nTAAAATklEQVR4nO3PUQkAIBTAwBfFaEYzqiH8OITBAtxmnf11wwUNaEEDWtCAFjSgBQ1oQQNa0IAWNKAFDWhBA1rQgBY0oAUNaEEDWtCAFjx2AXTMAIjyYikUAAAAAElFTkSuQmCC', 'base64');
  return zipSync(parts);
}
