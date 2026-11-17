import JSZip from 'jszip';

export interface ChartSeriesSpec {
  name: string;
  valuesRef: string; // e.g. "'4A. Counter Wise Summary'!$C$5:$C$11"
  colorHex: string; // 6-char RGB hex without #
}

export interface EmbeddedChartSpec {
  sheetName: string; // Exact worksheet name in workbook
  title: string;
  chartType: 'col' | 'bar' | 'pie';
  categoriesRef: string; // e.g. "'4A. Counter Wise Summary'!$B$5:$B$11"
  series: ChartSeriesSpec[];
  anchor: {
    fromCol: number;
    fromRow: number;
    toCol: number;
    toRow: number;
  };
}

function escapeXml(str: string): string {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export function colIndexToLetter(col1Based: number): string {
  let n = col1Based;
  let result = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result || 'A';
}

export function sheetRangeRef(
  sheetName: string,
  col1Based: number,
  startRow: number,
  endRow: number
): string {
  const colLetter = colIndexToLetter(col1Based);
  const safeSheet = `'${sheetName.replace(/'/g, "''")}'`;
  return `${safeSheet}!$${colLetter}$${startRow}:$${colLetter}$${endRow}`;
}

const PIE_SLICE_COLORS = [
  '1D4ED8',
  '0D9488',
  'D97706',
  '059669',
  '7C3AED',
  'DC2626',
  '0284C7',
  '64748B',
];

function buildChartXml(spec: EmbeddedChartSpec): string {
  const safeTitle = escapeXml(spec.title);

  if (spec.chartType === 'pie') {
    const firstSeries = spec.series[0];
    const sliceCount = 8;
    const dPtXml = Array.from({ length: sliceCount }, (_, idx) => {
      const color = PIE_SLICE_COLORS[idx % PIE_SLICE_COLORS.length];
      return `
        <c:dPt>
          <c:idx val="${idx}"/>
          <c:spPr>
            <a:solidFill><a:srgbClr val="${color}"/></a:solidFill>
            <a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
          </c:spPr>
        </c:dPt>`;
    }).join('');

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr><a:defRPr sz="1200" b="1"><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill></a:defRPr></a:pPr>
            <a:r><a:rPr lang="en-US" sz="1200" b="1"/><a:t>${safeTitle}</a:t></a:r>
          </a:p>
        </c:rich>
      </c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:plotArea>
      <c:layout/>
      <c:pieChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:v>${escapeXml(firstSeries.name)}</c:v></c:tx>
          ${dPtXml}
          <c:dLbls>
            <c:showLegendKey val="0"/>
            <c:showVal val="1"/>
            <c:showCatName val="0"/>
            <c:showSerName val="0"/>
            <c:showPercent val="1"/>
          </c:dLbls>
          <c:cat><c:strRef><c:f>${escapeXml(spec.categoriesRef)}</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>${escapeXml(firstSeries.valuesRef)}</c:f></c:numRef></c:val>
        </c:ser>
        <c:firstSliceAng val="0"/>
      </c:pieChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
  <c:spPr>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    <a:ln w="9525"><a:solidFill><a:srgbClr val="CBD5E1"/></a:solidFill></a:ln>
  </c:spPr>
</c:chartSpace>`;
  }

  const barDir = spec.chartType === 'bar' ? 'bar' : 'col';
  const seriesXml = spec.series
    .map(
      (s, idx) => `
      <c:ser>
        <c:idx val="${idx}"/>
        <c:order val="${idx}"/>
        <c:tx><c:v>${escapeXml(s.name)}</c:v></c:tx>
        <c:spPr>
          <a:solidFill><a:srgbClr val="${s.colorHex}"/></a:solidFill>
        </c:spPr>
        <c:dLbls>
          <c:showLegendKey val="0"/>
          <c:showVal val="1"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
        </c:dLbls>
        <c:cat><c:strRef><c:f>${escapeXml(spec.categoriesRef)}</c:f></c:strRef></c:cat>
        <c:val><c:numRef><c:f>${escapeXml(s.valuesRef)}</c:f></c:numRef></c:val>
      </c:ser>`
    )
    .join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr><a:defRPr sz="1200" b="1"><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill></a:defRPr></a:pPr>
            <a:r><a:rPr lang="en-US" sz="1200" b="1"/><a:t>${safeTitle}</a:t></a:r>
          </a:p>
        </c:rich>
      </c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="${barDir}"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        ${seriesXml}
        <c:gapWidth val="120"/>
        <c:axId val="1001"/>
        <c:axId val="1002"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1001"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="${barDir === 'bar' ? 'l' : 'b'}"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1002"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="1002"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="${barDir === 'bar' ? 'b' : 'l'}"/>
        <c:majorGridlines>
          <c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:ln></c:spPr>
        </c:majorGridlines>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1001"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
  <c:spPr>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    <a:ln w="9525"><a:solidFill><a:srgbClr val="CBD5E1"/></a:solidFill></a:ln>
  </c:spPr>
</c:chartSpace>`;
}

function buildDrawingXml(charts: Array<{ spec: EmbeddedChartSpec; chartRelId: string }>): string {
  const anchors = charts
    .map(
      ({ spec, chartRelId }, idx) => `
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>${spec.anchor.fromCol}</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>${spec.anchor.fromRow}</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>${spec.anchor.toCol}</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>${spec.anchor.toRow}</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="${idx + 2}" name="Chart ${idx + 1}"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                   r:id="${chartRelId}"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>`
    )
    .join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  ${anchors}
</xdr:wsDr>`;
}

export interface InternalHyperlinkSpec {
  sheetName: string; // Source worksheet name
  cellRef: string; // Source cell address e.g. "B5"
  location: string; // Target sheet & cell e.g. "'1. DSE Today Followup List'!A5"
  display?: string;
  tooltip?: string;
}

export async function embedNativeChartsInXlsx(
  xlsxBuffer: Buffer,
  charts: EmbeddedChartSpec[],
  hyperlinks: InternalHyperlinkSpec[] = []
): Promise<Buffer> {
  if (charts.length === 0 && hyperlinks.length === 0) return xlsxBuffer;

  const zip = await JSZip.loadAsync(xlsxBuffer);
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const workbookRelsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  const contentTypesXml = await zip.file('[Content_Types].xml')?.async('string');

  if (!workbookXml || !workbookRelsXml || !contentTypesXml) {
    return xlsxBuffer;
  }

  // Map sheet name -> rId -> xl/worksheets/sheetN.xml (attribute-order independent)
  const sheetRIdByName = new Map<string, string>();
  const sheetTagRegex = /<sheet\b([^>]*)\/?>/g;
  let m: RegExpExecArray | null;
  while ((m = sheetTagRegex.exec(workbookXml)) !== null) {
    const attrs = m[1];
    const nameMatch = attrs.match(/\bname="([^"]+)"/);
    const ridMatch = attrs.match(/\br:id="([^"]+)"/);
    if (!nameMatch || !ridMatch) continue;
    const rawName = nameMatch[1]
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");
    sheetRIdByName.set(rawName, ridMatch[1]);
  }

  const targetByRId = new Map<string, string>();
  const relTagRegex = /<Relationship\b([^>]*)\/?>/g;
  while ((m = relTagRegex.exec(workbookRelsXml)) !== null) {
    const attrs = m[1];
    const idMatch = attrs.match(/\bId="([^"]+)"/);
    const targetMatch = attrs.match(/\bTarget="([^"]+)"/);
    if (idMatch && targetMatch) {
      targetByRId.set(idMatch[1], targetMatch[1]);
    }
  }

  // Group charts by sheetName
  const chartsBySheet = new Map<string, EmbeddedChartSpec[]>();
  for (const ch of charts) {
    const list = chartsBySheet.get(ch.sheetName) || [];
    list.push(ch);
    chartsBySheet.set(ch.sheetName, list);
  }

  let chartCounter = 1;
  let drawingCounter = 1;
  const newContentTypeOverrides: string[] = [];

  for (const [sheetName, sheetCharts] of chartsBySheet.entries()) {
    const rId = sheetRIdByName.get(sheetName);
    if (!rId) continue;
    const relTarget = targetByRId.get(rId);
    if (!relTarget) continue;

    const normalizedTarget = relTarget.replace(/^\/?/, '');
    const sheetPath = normalizedTarget.startsWith('xl/')
      ? normalizedTarget
      : `xl/${normalizedTarget}`;
    const sheetXml = await zip.file(sheetPath)?.async('string');
    if (!sheetXml) continue;

    const sheetBaseName = sheetPath.split('/').pop()!;
    const sheetRelsPath = `xl/worksheets/_rels/${sheetBaseName}.rels`;
    const existingSheetRels = await zip.file(sheetRelsPath)?.async('string');

    const drawingIndex = drawingCounter++;
    const drawingFileName = `drawing${drawingIndex}.xml`;
    const drawingZipPath = `xl/drawings/${drawingFileName}`;
    const drawingRelsZipPath = `xl/drawings/_rels/${drawingFileName}.rels`;

    const drawingChartEntries: Array<{ spec: EmbeddedChartSpec; chartRelId: string }> = [];
    const drawingRelLines: string[] = [];

    for (let i = 0; i < sheetCharts.length; i++) {
      const cIdx = chartCounter++;
      const chartFileName = `chart${cIdx}.xml`;
      const chartZipPath = `xl/charts/${chartFileName}`;
      const relId = `rId${i + 1}`;

      zip.file(chartZipPath, buildChartXml(sheetCharts[i]));
      newContentTypeOverrides.push(
        `<Override PartName="/${chartZipPath}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`
      );
      drawingRelLines.push(
        `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/${chartFileName}"/>`
      );
      drawingChartEntries.push({ spec: sheetCharts[i], chartRelId: relId });
    }

    zip.file(drawingZipPath, buildDrawingXml(drawingChartEntries));
    zip.file(
      drawingRelsZipPath,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${drawingRelLines.join('\n  ')}
</Relationships>`
    );
    newContentTypeOverrides.push(
      `<Override PartName="/${drawingZipPath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`
    );

    // Attach drawing to sheet rels
    let drawingRelId = 'rId1';
    if (existingSheetRels) {
      const existingIds = Array.from(existingSheetRels.matchAll(/Id="rId(\d+)"/g)).map((x) =>
        Number(x[1])
      );
      const nextNum = existingIds.length > 0 ? Math.max(...existingIds) + 1 : 1;
      drawingRelId = `rId${nextNum}`;
      const updatedRels = existingSheetRels.replace(
        '</Relationships>',
        `  <Relationship Id="${drawingRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/${drawingFileName}"/>\n</Relationships>`
      );
      zip.file(sheetRelsPath, updatedRels);
    } else {
      zip.file(
        sheetRelsPath,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${drawingRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/${drawingFileName}"/>
</Relationships>`
      );
    }

    // Insert <drawing r:id="..."/> in sheetXml before <tableParts>, <extLst>, or </worksheet>
    const drawingTag = `<drawing r:id="${drawingRelId}"/>`;
    let updatedSheetXml = sheetXml;
    if (updatedSheetXml.includes('<tableParts')) {
      updatedSheetXml = updatedSheetXml.replace('<tableParts', `${drawingTag}<tableParts`);
    } else if (updatedSheetXml.includes('<extLst')) {
      updatedSheetXml = updatedSheetXml.replace('<extLst', `${drawingTag}<extLst`);
    } else {
      updatedSheetXml = updatedSheetXml.replace('</worksheet>', `${drawingTag}</worksheet>`);
    }
    zip.file(sheetPath, updatedSheetXml);
  }

  if (newContentTypeOverrides.length > 0) {
    const updatedTypes = contentTypesXml.replace(
      '</Types>',
      `  ${newContentTypeOverrides.join('\n  ')}\n</Types>`
    );
    zip.file('[Content_Types].xml', updatedTypes);
  }

  // Group internal drill-down hyperlinks by sheetName
  if (hyperlinks.length > 0) {
    const linksBySheet = new Map<string, InternalHyperlinkSpec[]>();
    for (const link of hyperlinks) {
      const list = linksBySheet.get(link.sheetName) || [];
      list.push(link);
      linksBySheet.set(link.sheetName, list);
    }

    for (const [sheetName, sheetLinks] of linksBySheet.entries()) {
      const rId = sheetRIdByName.get(sheetName);
      if (!rId) continue;
      const relTarget = targetByRId.get(rId);
      if (!relTarget) continue;
      const normalizedTarget = relTarget.replace(/^\/?/, '');
      const sheetPath = normalizedTarget.startsWith('xl/')
        ? normalizedTarget
        : `xl/${normalizedTarget}`;
      const sheetXml = await zip.file(sheetPath)?.async('string');
      if (!sheetXml) continue;

      const linkTags = sheetLinks
        .map((l) => {
          const attrs = [
            `ref="${escapeXml(l.cellRef)}"`,
            `location="${escapeXml(l.location)}"`,
          ];
          if (l.display) attrs.push(`display="${escapeXml(l.display)}"`);
          if (l.tooltip) attrs.push(`tooltip="${escapeXml(l.tooltip)}"`);
          return `<hyperlink ${attrs.join(' ')}/>`;
        })
        .join('');

      let updatedXml = sheetXml;
      if (updatedXml.includes('</hyperlinks>')) {
        updatedXml = updatedXml.replace('</hyperlinks>', `${linkTags}</hyperlinks>`);
      } else {
        const block = `<hyperlinks>${linkTags}</hyperlinks>`;
        const anchors = [
          '<printOptions',
          '<pageMargins',
          '<pageSetup',
          '<headerFooter',
          '<rowBreaks',
          '<drawing',
          '<tableParts',
          '<extLst',
          '</worksheet>',
        ];
        const targetTag = anchors.find((tag) => updatedXml.includes(tag)) || '</worksheet>';
        updatedXml = updatedXml.replace(targetTag, `${block}${targetTag}`);
      }
      zip.file(sheetPath, updatedXml);
    }
  }

  const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  return Buffer.from(out);
}

