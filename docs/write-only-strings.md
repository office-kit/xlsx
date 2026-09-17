# Write-only string storage

`createWriteOnlyWorkbook` retains a bounded shared-string table across all
worksheets. It needs no new option: both plain strings and rich text follow
the same policy. The normal workbook model and `saveWorkbook` keep their
existing shared-string behavior.

## Admission and reuse

- An existing value always reuses its shared-string ID, even after the table
  stops admitting entries or a new sheet starts. IDs are never evicted.
- A new value enters the table only if the resulting table has at most
  100,000 entries and at most 8 MiB of accounted payload.
- The first new value that cannot fit is emitted as `t="inlineStr"`. From
  then on, all new values are inline, even if a later smaller value could fit.
  This makes the transition predictable and avoids repeated admission work.
- Plain strings are keyed by their original text. Rich text is keyed by its
  serialized runs, including formatting, and snapshotted when appended.
  Later mutations of caller-owned runs or fonts cannot alter earlier cells.

Payload accounting charges two bytes per UTF-16 code unit in both the lookup
key and the serialized `<t>`/`<r>` content. It conservatively charges both
even when they share a JavaScript string. XML escapes and rich-text formatting
therefore count toward the limit. Map/array/object overhead is bounded by the
entry cap but is not included in the 8 MiB budget; engine allocation details
and temporary values mean actual heap usage is higher.

## Serialization and compatibility

Shared and inline cells use the same text escaping, whitespace preservation,
and rich-text formatting. Cell styles are independent of string storage.
Both storage forms are valid SpreadsheetML and can coexist within a sheet.
The shared-string part, relationship, and content-type override are emitted
only when the table contains an entry. Finalization streams the shared-string
XML in chunks of at most 16,384 UTF-16 code units without splitting surrogate
pairs, instead of building another complete table-sized XML/byte buffer.

Public method signatures and cell values are unchanged. ZIP bytes, string IDs,
and the shared/inline representation may differ from earlier releases; code
that inspects raw XML must accept both forms. With the same row order, options,
and pinned `mtime`, output remains deterministic within this implementation.
Streaming reads preserve text values but flatten rich-text formatting, as they
do for shared strings; model-based reads preserve the runs.

## Memory scope and tradeoffs

This bounds retained string state, not the entire workbook's memory. Styles,
sheet metadata, column settings, the current row/value, compression scratch,
and the destination's output buffers consume additional memory. A single
very large value still needs temporary serialization space. Use a streaming
sink and bounded style/sheet counts for large exports.

Repeated values admitted before the cap retain shared-string compression.
New values after the cap are repeated inline if they recur, which can increase
XML size and serialization work. ZIP compression still applies. There is no
unbounded mode or tuning option; the bounded policy is the default write-only
contract.
