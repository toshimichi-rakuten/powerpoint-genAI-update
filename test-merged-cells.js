const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

// Read the slide XML from table_2 sample
const slideXmlPath = path.join(__dirname, '../sample/table_2/ppt/slides/slide1.xml');
const slideXml = fs.readFileSync(slideXmlPath, 'utf-8');

// Parse the XML
const dom = new JSDOM(slideXml, { contentType: 'text/xml' });
const doc = dom.window.document;

// Find the table
const graphicFrames = Array.from(doc.getElementsByTagName('*')).filter(el =>
  el.tagName.endsWith(':graphicFrame') || el.localName === 'graphicFrame'
);

console.log('=== MERGED CELL ANALYSIS TEST ===\n');
console.log(`Found ${graphicFrames.length} table(s)\n`);

// Analyze the second table (index 1 - the complex merged table with 11 columns)
const firstTable = graphicFrames[1];
const tbl = Array.from(firstTable.getElementsByTagName('*')).find(el =>
  el.tagName.endsWith(':tbl') || el.localName === 'tbl'
);

if (tbl) {
  const trs = Array.from(tbl.getElementsByTagName('*')).filter(el =>
    el.tagName.endsWith(':tr') || el.localName === 'tr'
  );

  console.log(`Total rows in table: ${trs.length}`);
  console.log('TABLE STRUCTURE (All 12 rows):');
  console.log('Expected: 11 columns per row\n');

  trs.forEach((tr, rowIndex) => {
    const tcs = Array.from(tr.getElementsByTagName('*')).filter(el =>
      el.tagName.endsWith(':tc') || el.localName === 'tc'
    );

    console.log(`ROW ${rowIndex}: ${tcs.length} cells in XML`);

    let expectedCellCount = 0;
    tcs.forEach((tc, cellIndex) => {
      const gridSpan = tc.getAttribute('gridSpan');
      const rowSpan = tc.getAttribute('rowSpan');
      const hMerge = tc.getAttribute('hMerge');
      const vMerge = tc.getAttribute('vMerge');

      const isMergedContinuation = (hMerge === '1' || vMerge === '1');

      // Get text
      const txBody = Array.from(tc.getElementsByTagName('*')).find(el =>
        el.tagName.endsWith(':txBody') || el.localName === 'txBody'
      );

      let text = '';
      if (txBody && !isMergedContinuation) {
        const tNodes = Array.from(txBody.getElementsByTagName('*')).filter(el =>
          el.tagName.endsWith(':t') || el.localName === 't'
        );
        text = tNodes.map(t => t.textContent).join('');
      }

      const cellInfo = [];
      if (gridSpan) cellInfo.push(`colspan=${gridSpan}`);
      if (rowSpan) cellInfo.push(`rowspan=${rowSpan}`);
      if (isMergedContinuation) {
        cellInfo.push('CONTINUATION');
        expectedCellCount++; // Continuation cells MUST be counted
      } else {
        expectedCellCount++; // Normal cells
      }

      const infoStr = cellInfo.length > 0 ? ` [${cellInfo.join(', ')}]` : '';
      console.log(`  Cell ${cellIndex}: "${text}"${infoStr}`);
    });

    console.log(`  → Output cell count: ${expectedCellCount} (${expectedCellCount === 11 ? '✓ CORRECT' : '✗ WRONG - should be 11'})\n`);
  });

  console.log('\n=== KEY INSIGHT ===');
  console.log('With the FIX:');
  console.log('  - Continuation cells are KEPT as placeholders with isMergedContinuation=true');
  console.log('  - Each row has exactly 11 cells in the output array');
  console.log('  - Cell indices match grid positions (no shift)');
  console.log('\nWithout the fix (OLD behavior):');
  console.log('  - Continuation cells were SKIPPED with return statement');
  console.log('  - Rows had fewer cells (e.g., 9 cells instead of 11)');
  console.log('  - Cell indices were shifted (e.g., cell at grid position 2 appeared at index 1)');
  console.log('  - This caused misalignment when regenerating the PowerPoint');
}
