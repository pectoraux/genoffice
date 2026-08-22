/**
 * Pivot fixture builder — creates a minimal XLSX with a pivot table
 * for testing workbook:read-pivot-definition.
 *
 * The fixture contains:
 *   - A data sheet (Sheet1) with 4 rows of data (Name, Category, Value)
 *   - A pivot cache definition (xl/pivotCache/pivotCacheDefinition1.xml)
 *   - A pivot table definition (xl/pivotTables/pivotTable1.xml)
 *   - A pivot table worksheet (Sheet2) where the pivot output lives
 *   - All required content-types, relationships, and workbook.xml entries
 *
 * This is the smallest valid XLSX with a pivot table that the Rust
 * sidecar can open and the xlsx-gateway parser can parse.
 */
import JSZip from 'jszip'

export async function buildPivotFixture(): Promise<Buffer> {
  const zip = new JSZip()

  // Content types — must declare pivot table + cache content types
  zip.file('[Content_Types].xml', pivotContentTypes)
  zip.file('_rels/.rels', packageRelationships)

  // Workbook with 2 sheets: Data (Sheet1) + Pivot (Sheet2)
  zip.file('xl/workbook.xml', pivotWorkbook)
  zip.file('xl/_rels/workbook.xml.rels', pivotWorkbookRelationships)

  // Data sheet (Sheet1) — source data for the pivot
  zip.file('xl/worksheets/sheet1.xml', dataWorksheet)

  // Pivot sheet (Sheet2) — where the pivot table output lives
  zip.file('xl/worksheets/sheet2.xml', pivotWorksheet)

  // Pivot cache definition
  zip.file('xl/pivotCache/pivotCacheDefinition1.xml', pivotCacheDefinition)
  zip.file('xl/_rels/worksheets/sheet2.xml.rels', sheet2Relationships)

  // Pivot table definition
  zip.file('xl/pivotTables/pivotTable1.xml', pivotTableDefinition)

  // Styles (minimal)
  zip.file('xl/styles.xml', minimalStyles)

  // Shared strings (minimal — the data uses inline strings)
  zip.file('xl/sharedStrings.xml', minimalSharedStrings)

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

const pivotContentTypes = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
</Types>`

const packageRelationships = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

const pivotWorkbook = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="Pivot" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`

const pivotWorkbookRelationships = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`

const dataWorksheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Category</t></is></c><c r="C1" t="inlineStr"><is><t>Value</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Apple</t></is></c><c r="B2" t="inlineStr"><is><t>Fruit</t></is></c><c r="C2"><v>10</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>Banana</t></is></c><c r="B3" t="inlineStr"><is><t>Fruit</t></is></c><c r="C3"><v>20</v></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>Carrot</t></is></c><c r="B4" t="inlineStr"><is><t>Vegetable</t></is></c><c r="C4"><v>15</v></c></row>
  </sheetData>
</worksheet>`

const pivotWorksheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>`

const pivotCacheDefinition = `<?xml version="1.0" encoding="UTF-8"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshOnLoad="1" refreshedBy="Test" createdVersion="3" refreshedVersion="3" minRefreshableVersion="3" recordCount="3">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:C4" sheet="Data"/>
  </cacheSource>
  <cacheFields count="3">
    <cacheField name="Name" numFmtId="0">
      <sharedItems count="3"><s v="Apple"/><s v="Banana"/><s v="Carrot"/></sharedItems>
    </cacheField>
    <cacheField name="Category" numFmtId="0">
      <sharedItems count="2"><s v="Fruit"/><s v="Vegetable"/></sharedItems>
    </cacheField>
    <cacheField name="Value" numFmtId="0">
      <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="10" maxValue="20" count="3"><n v="10"/><n v="20"/><n v="15"/></sharedItems>
    </cacheField>
  </cacheFields>
</pivotCacheDefinition>`

const pivotTableDefinition = `<?xml version="1.0" encoding="UTF-8"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="3" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="3" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">
  <location ref="A3:B7" firstHeaderRow="1" firstDataRow="2" firstDataCol="1" rowPageCount="1" colPageCount="1"/>
  <pivotFields count="3">
    <pivotField axis="axisRow" showAll="0">
      <items count="4"><item x="0"/><item x="1"/><item x="2"/><item t="default"/></items>
    </pivotField>
    <pivotField axis="axisCol" showAll="0">
      <items count="3"><item x="0"/><item x="1"/><item t="default"/></items>
    </pivotField>
    <pivotField dataField="1" showAll="0"/>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <colFields count="1"><field x="1"/></colFields>
  <dataFields count="1" dataField="2" caption="Sum of Value">
    <dataField name="Sum of Value" fld="2" baseField="0" baseItem="0"/>
  </dataFields>
  <pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>
</pivotTableDefinition>`

const sheet2Relationships = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>`

const minimalStyles = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs>
</styleSheet>`

const minimalSharedStrings = `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>`
