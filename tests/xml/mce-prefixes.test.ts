import { describe, expect, it } from 'vitest';
import { parseXml } from '../../src/xml/parser.js';
import { serializeXml } from '../../src/xml/serializer.js';

// `Requires` / `mc:Ignorable` name prefixes, so the serializer has to keep the
// prefixes Excel uses or those references end up unbound.
describe('Excel extension prefixes', () => {
  it('keeps a14 bound for an mc:Choice that requires it', () => {
    const src =
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">' +
      '<a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}"><a14:compatExt spid="_x0000_s1025"/></a:ext></mc:Choice></mc:AlternateContent></xdr:wsDr>';
    const out = new TextDecoder().decode(serializeXml(parseXml(src)));
    expect(out).toContain('Requires="a14"');
    expect(out).toContain('xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"');
    expect(out).toContain('<a14:compatExt');
  });

  it('writes the revision namespace as xr, as Excel does', () => {
    const src = '<a xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xr:uid="{1}"/>';
    expect(new TextDecoder().decode(serializeXml(parseXml(src)))).toContain('xr:uid="{1}"');
  });
});
