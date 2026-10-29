import JSZip from 'jszip';

export type NativeExcelChartType = 'COLUMN' | 'BAR' | 'LINE' | 'DOUGHNUT';

export interface NativeChartSeriesSpec {
  name: string;
  nameFormula: string; // e.g. "'Dashboard Data'!$B$4"
  catFormula: string; // e.g. "'Dashboard Data'!$A$5:$A$16"
  valFormula: string; // e.g. "'Dashboard Data'!$B$5:$B$16"
  colorHex: string; // 6-char hex without #, e.g. "2A78D6"
  categories: string[];
  values: number[];
  sliceColorsHex?: string[]; // Used for DOUGHNUT slices
}

export interface NativeExcelChartSpec {
  sheetName: string; // Target worksheet where chart is embedded
  title: string;
  chartType: NativeExcelChartType;
  fromCol: number; // 0-based col index
  fromRow: number; // 0-based row index
  toCol: number; // 0-based col index
  toRow: number; // 0-based row index
  series: NativeChartSeriesSpec[];
}

function escapeXml(str: string): string {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export function buildChartXml(spec: NativeExcelChartSpec): string {
  const isDoughnut = spec.chartType === 'DOUGHNUT';
  const isLine = spec.chartType === 'LINE';
  const isHorizontalBar = spec.chartType === 'BAR';

  const seriesXml = spec.series
    .map((s, idx) => {
      const catPts = s.categories
        .map(
          (cat, i) =>
            `<c:pt idx="${i}"><c:v>${escapeXml(cat)}</c:v></c:pt>`
        )
        .join('');

      const valPts = s.values
        .map((val, i) => {
          const safeNum = Number.isFinite(val) ? Number(val.toFixed(2)) : 0;
          return `<c:pt idx="${i}"><c:v>${safeNum}</c:v></c:pt>`;
        })
        .join('');

      const dPtsXml =
        isDoughnut && s.sliceColorsHex && s.sliceColorsHex.length > 0
          ? s.categories
              .map((_, sliceIdx) => {
                const hex =
                  s.sliceColorsHex![sliceIdx % s.sliceColorsHex!.length] ||
                  '2A78D6';
                return `<c:dPt><c:idx val="${sliceIdx}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`;
              })
              .join('')
          : '';

      const spPrXml = isDoughnut
        ? ''
        : isLine
          ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${s.colorHex}"/></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="${s.colorHex}"/></a:solidFill></c:spPr></c:marker>`
          : `<c:spPr><a:solidFill><a:srgbClr val="${s.colorHex}"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="${s.colorHex}"/></a:solidFill></a:ln></c:spPr><c:invertIfNegative val="0"/>`;

      const smoothXml = isLine ? '<c:smooth val="0"/>' : '';

      return `<c:ser>
        <c:idx val="${idx}"/>
        <c:order val="${idx}"/>
        <c:tx>
          <c:strRef>
            <c:f>${escapeXml(s.nameFormula)}</c:f>
            <c:strCache>
              <c:ptCount val="1"/>
              <c:pt idx="0"><c:v>${escapeXml(s.name)}</c:v></c:pt>
            </c:strCache>
          </c:strRef>
        </c:tx>
        ${spPrXml}
        ${dPtsXml}
        <c:cat>
          <c:strRef>
            <c:f>${escapeXml(s.catFormula)}</c:f>
            <c:strCache>
              <c:ptCount val="${s.categories.length}"/>
              ${catPts}
            </c:strCache>
          </c:strRef>
        </c:cat>
        <c:val>
          <c:numRef>
            <c:f>${escapeXml(s.valFormula)}</c:f>
            <c:numCache>
              <c:formatCode>#,##0</c:formatCode>
              <c:ptCount val="${s.values.length}"/>
              ${valPts}
            </c:numCache>
          </c:numRef>
        </c:val>
        ${smoothXml}
      </c:ser>`;
    })
    .join('\n');

  let plotTypeXml = '';
  if (isDoughnut) {
    plotTypeXml = `<c:doughnutChart>
      <c:varyColors val="1"/>
      ${seriesXml}
      <c:dLbls>
        <c:showLegendKey val="0"/>
        <c:showVal val="0"/>
        <c:showCatName val="0"/>
        <c:showSerName val="0"/>
        <c:showPercent val="1"/>
        <c:showBubbleSize val="0"/>
        <c:showLeaderLines val="1"/>
      </c:dLbls>
      <c:firstSliceAng val="0"/>
      <c:holeSize val="55"/>
    </c:doughnutChart>`;
  } else if (isLine) {
    plotTypeXml = `<c:lineChart>
      <c:grouping val="standard"/>
      <c:varyColors val="0"/>
      ${seriesXml}
      <c:dLbls>
        <c:showLegendKey val="0"/>
        <c:showVal val="0"/>
        <c:showCatName val="0"/>
        <c:showSerName val="0"/>
        <c:showPercent val="0"/>
        <c:showBubbleSize val="0"/>
      </c:dLbls>
      <c:marker val="1"/>
      <c:smooth val="0"/>
      <c:axId val="1001"/>
      <c:axId val="1002"/>
    </c:lineChart>`;
  } else {
    const barDir = isHorizontalBar ? 'bar' : 'col';
    plotTypeXml = `<c:barChart>
      <c:barDir val="${barDir}"/>
      <c:grouping val="clustered"/>
      <c:varyColors val="0"/>
      ${seriesXml}
      <c:dLbls>
        <c:showLegendKey val="0"/>
        <c:showVal val="0"/>
        <c:showCatName val="0"/>
        <c:showSerName val="0"/>
        <c:showPercent val="0"/>
        <c:showBubbleSize val="0"/>
      </c:dLbls>
      <c:gapWidth val="160"/>
      <c:overlap val="-15"/>
      <c:axId val="1001"/>
      <c:axId val="1002"/>
    </c:barChart>`;
  }

  const axesXml = isDoughnut
    ? ''
    : `<c:catAx>
        <c:axId val="1001"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="${isHorizontalBar ? 'l' : 'b'}"/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:spPr>
          <a:ln w="9525"><a:solidFill><a:srgbClr val="CBD5E1"/></a:solidFill></a:ln>
        </c:spPr>
        <c:txPr>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr><a:defRPr sz="850" b="0"><a:solidFill><a:srgbClr val="1F2937"/></a:solidFill></a:defRPr></a:pPr>
            <a:endParaRPr lang="en-US"/>
          </a:p>
        </c:txPr>
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
        <c:axPos val="${isHorizontalBar ? 'b' : 'l'}"/>
        <c:majorGridlines>
          <c:spPr>
            <a:ln w="9525"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:ln>
          </c:spPr>
        </c:majorGridlines>
        <c:numFmt formatCode="#,##0" sourceLinked="0"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:spPr><a:ln><a:noFill/></a:ln></c:spPr>
        <c:txPr>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr><a:defRPr sz="850" b="0"><a:solidFill><a:srgbClr val="475569"/></a:solidFill></a:defRPr></a:pPr>
            <a:endParaRPr lang="en-US"/>
          </a:p>
        </c:txPr>
        <c:crossAx val="1001"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr>
              <a:defRPr sz="1050" b="1">
                <a:solidFill><a:srgbClr val="0B1F3A"/></a:solidFill>
              </a:defRPr>
            </a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="1050" b="1">
                <a:solidFill><a:srgbClr val="0B1F3A"/></a:solidFill>
              </a:rPr>
              <a:t>${escapeXml(spec.title)}</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      ${plotTypeXml}
      ${axesXml}
      <c:spPr>
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
      </c:spPr>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
      <c:overlay val="0"/>
      <c:txPr>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:pPr><a:defRPr sz="850" b="0"><a:solidFill><a:srgbClr val="1F2937"/></a:solidFill></a:defRPr></a:pPr>
          <a:endParaRPr lang="en-US"/>
        </a:p>
      </c:txPr>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
  <c:spPr>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    <a:ln w="12700">
      <a:solidFill><a:srgbClr val="CBD5E1"/></a:solidFill>
    </a:ln>
  </c:spPr>
</c:chartSpace>`;
}

export function buildDrawingXml(
  chartsForSheet: Array<{ spec: NativeExcelChartSpec; relId: string; drawingObjId: number }>
): string {
  const anchors = chartsForSheet
    .map(
      ({ spec, relId, drawingObjId }) => `<xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>${spec.fromCol}</xdr:col>
      <xdr:colOff>114300</xdr:colOff>
      <xdr:row>${spec.fromRow}</xdr:row>
      <xdr:rowOff>76200</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>${spec.toCol}</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>${spec.toRow}</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="${drawingObjId}" name="Chart ${drawingObjId}: ${escapeXml(spec.title)}"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${relId}"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>`
    )
    .join('\n');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
${anchors}
</xdr:wsDr>`;
}

/**
 * Injects native ECMA-376 OpenXML chart objects and drawing anchors into an ExcelJS-generated XLSX buffer.
 */
export async function injectNativeChartsIntoXlsx(
  xlsxBuffer: ArrayBuffer | Uint8Array,
  chartSpecs: NativeExcelChartSpec[]
): Promise<Uint8Array> {
  if (chartSpecs.length === 0) {
    return xlsxBuffer instanceof Uint8Array ? xlsxBuffer : new Uint8Array(xlsxBuffer);
  }

  const zip = await JSZip.loadAsync(xlsxBuffer);

  // 1. Parse xl/workbook.xml and xl/_rels/workbook.xml.rels to map sheetName -> sheetFile (e.g. "worksheets/sheet2.xml")
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const workbookRelsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  if (!workbookXml || !workbookRelsXml) {
    return xlsxBuffer instanceof Uint8Array ? xlsxBuffer : new Uint8Array(xlsxBuffer);
  }

  const relIdToTarget = new Map<string, string>();
  const relRegex = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/?>/g;
  let m: RegExpExecArray | null;
  while ((m = relRegex.exec(workbookRelsXml)) !== null) {
    relIdToTarget.set(m[1], m[2]);
  }
  // Also match if Target appears before Id
  const relRegexAlt = /<Relationship\b[^>]*\bTarget="([^"]+)"[^>]*\bId="([^"]+)"[^>]*\/?>/g;
  while ((m = relRegexAlt.exec(workbookRelsXml)) !== null) {
    relIdToTarget.set(m[2], m[1]);
  }

  const sheetNameToTarget = new Map<string, string>();
  const sheetRegex = /<sheet\b[^>]*\bname="([^"]+)"[^>]*\br:id="([^"]+)"[^>]*\/?>/g;
  while ((m = sheetRegex.exec(workbookXml)) !== null) {
    const sName = m[1]
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");
    const rId = m[2];
    const target = relIdToTarget.get(rId);
    if (target) {
      sheetNameToTarget.set(sName, target.startsWith('xl/') ? target : `xl/${target.replace(/^\//, '')}`);
    }
  }

  // Group chartSpecs by sheetName
  const chartsBySheet = new Map<string, NativeExcelChartSpec[]>();
  for (const spec of chartSpecs) {
    if (!sheetNameToTarget.has(spec.sheetName)) continue;
    const arr = chartsBySheet.get(spec.sheetName) || [];
    arr.push(spec);
    chartsBySheet.set(spec.sheetName, arr);
  }

  let globalChartCounter = 1;
  let globalDrawingCounter = 1;
  const createdChartFiles: string[] = [];
  const createdDrawingFiles: string[] = [];

  for (const [sheetName, specs] of chartsBySheet.entries()) {
    const sheetPath = sheetNameToTarget.get(sheetName)!; // e.g. "xl/worksheets/sheet2.xml"
    const sheetFilename = sheetPath.split('/').pop()!; // e.g. "sheet2.xml"
    const sheetRelsPath = `xl/worksheets/_rels/${sheetFilename}.rels`;

    const drawingIndex = globalDrawingCounter++;
    const drawingPath = `xl/drawings/drawing${drawingIndex}.xml`;
    const drawingRelsPath = `xl/drawings/_rels/drawing${drawingIndex}.xml.rels`;
    createdDrawingFiles.push(`/${drawingPath}`);

    const drawingCharts: Array<{ spec: NativeExcelChartSpec; relId: string; drawingObjId: number }> = [];
    const drawingRelEntries: string[] = [];

    specs.forEach((spec, idx) => {
      const chartIndex = globalChartCounter++;
      const chartPath = `xl/charts/chart${chartIndex}.xml`;
      createdChartFiles.push(`/${chartPath}`);

      const chartRelId = `rIdChart${idx + 1}`;
      drawingCharts.push({
        spec,
        relId: chartRelId,
        drawingObjId: idx + 2,
      });

      drawingRelEntries.push(
        `<Relationship Id="${chartRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartIndex}.xml"/>`
      );

      zip.file(chartPath, buildChartXml(spec));
    });

    // Write drawing XML and drawing.xml.rels
    zip.file(drawingPath, buildDrawingXml(drawingCharts));
    const drawingRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${drawingRelEntries.join('\n  ')}
</Relationships>`;
    zip.file(drawingRelsPath, drawingRelsXml);

    // Link sheet{S}.xml to drawing{M}.xml
    const drawingSheetRelId = `rIdGspuDrawing${drawingIndex}`;
    const existingSheetRels = await zip.file(sheetRelsPath)?.async('string');
    const drawingRelTag = `<Relationship Id="${drawingSheetRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingIndex}.xml"/>`;

    if (existingSheetRels) {
      const updatedSheetRels = existingSheetRels.replace(
        /<\/Relationships>/,
        `  ${drawingRelTag}\n</Relationships>`
      );
      zip.file(sheetRelsPath, updatedSheetRels);
    } else {
      const newSheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${drawingRelTag}
</Relationships>`;
      zip.file(sheetRelsPath, newSheetRels);
    }

    // Inject <drawing r:id="..."/> into sheet{S}.xml before <tableParts> or </worksheet>
    const sheetXml = await zip.file(sheetPath)?.async('string');
    if (sheetXml) {
      let updatedSheetXml = sheetXml;
      if (!updatedSheetXml.includes('xmlns:r=')) {
        updatedSheetXml = updatedSheetXml.replace(
          '<worksheet ',
          '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        );
      }
      const drawingTag = `<drawing r:id="${drawingSheetRelId}"/>`;
      if (updatedSheetXml.includes('<tableParts')) {
        updatedSheetXml = updatedSheetXml.replace('<tableParts', `${drawingTag}<tableParts`);
      } else {
        updatedSheetXml = updatedSheetXml.replace('</worksheet>', `${drawingTag}</worksheet>`);
      }
      zip.file(sheetPath, updatedSheetXml);
    }
  }

  // 2. Register Content Types in [Content_Types].xml
  const contentTypesXml = await zip.file('[Content_Types].xml')?.async('string');
  if (contentTypesXml) {
    const overrides: string[] = [];
    for (const dPath of createdDrawingFiles) {
      if (!contentTypesXml.includes(`PartName="${dPath}"`)) {
        overrides.push(
          `<Override PartName="${dPath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`
        );
      }
    }
    for (const cPath of createdChartFiles) {
      if (!contentTypesXml.includes(`PartName="${cPath}"`)) {
        overrides.push(
          `<Override PartName="${cPath}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`
        );
      }
    }
    if (overrides.length > 0) {
      const updatedContentTypes = contentTypesXml.replace(
        /<\/Types>/,
        `  ${overrides.join('\n  ')}\n</Types>`
      );
      zip.file('[Content_Types].xml', updatedContentTypes);
    }
  }

  return await zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}
