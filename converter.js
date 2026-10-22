/* Local Excel reader and template-preserving exporter. No network or storage APIs. */
(function (root) {
  'use strict';
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const CHANNELS = {Shopee: {Columbia:'M01', Vans:'M02', Timberland:'M03'}, Lazada: {Columbia:'M04', Timberland:'M05', Vans:'M06'}, Zalora: {Timberland:'M16', Vans:'M19'}};
  const HEADERS = ['ChannelCode','eRetail SKU','Channel SKU','Channel Product Id','Back Order Inventory','Channel Price','Buffer Stock','Picking Instructions','Status','Mode'];
  const LIMIT = 5198;
  const elements = (node, name) => Array.from(node.getElementsByTagNameNS(NS, name));
  const normalize = value => String(value ?? '').trim().toLowerCase().replace(/\s+/g, '');
  function xml(text) {
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('The workbook contains invalid XML. Open it in Excel and save a fresh .xlsx copy.');
    return doc;
  }
  const escapeXml = value => String(value).replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[char]));
  function columnIndex(ref) {
    const letters = ref.match(/^[A-Z]+/i)?.[0];
    if (!letters) throw new Error('A worksheet has an invalid cell reference.');
    return [...letters.toUpperCase()].reduce((n,c) => n*26+c.charCodeAt(0)-64,0)-1;
  }
  async function openWorkbook(data) {
    let zip;
    try { zip = await JSZip.loadAsync(data); }
    catch { throw new Error('This file is not a supported Excel workbook. For legacy .xls or CSV files, open in Excel and Save As .xlsx first.'); }
    // Bound expanded content before reading XML to avoid accidental memory exhaustion.
    let expanded = 0;
    for (const entry of Object.values(zip.files)) {
      expanded += entry._data?.uncompressedSize || 0;
      if (expanded > 160*1024*1024) throw new Error('The expanded workbook exceeds 160 MB. Export a smaller file.');
    }
    const read = async path => {
      const file = zip.file(path);
      if (!file) throw new Error('Missing workbook component: '+path);
      return xml(await file.async('string'));
    };
    const workbook = await read('xl/workbook.xml');
    const relations = await read('xl/_rels/workbook.xml.rels');
    const rels = new Map(Array.from(relations.documentElement.children).filter(r => r.getAttribute('TargetMode') !== 'External').map(r => [r.getAttribute('Id'), r.getAttribute('Target')]));
    const shared = zip.file('xl/sharedStrings.xml') ? elements(await read('xl/sharedStrings.xml'),'si').map(item => elements(item,'t').map(t=>t.textContent).join('')) : [];
    const sheets = elements(workbook,'sheet').map(sheet => {
      const id = sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
      const target = rels.get(id);
      if (!target) return null;
      const path = new URL(target, 'https://local.invalid/xl/workbook.xml').pathname.slice(1);
      return {name:sheet.getAttribute('name'), path};
    }).filter(Boolean);
    if (!sheets.length) throw new Error('No supported worksheets were found.');
    async function rows(sheetIndex) {
      const sheet = sheets[sheetIndex];
      if (!sheet) throw new Error('Select a valid worksheet.');
      return elements(await read(sheet.path),'row').map(row => {
        const cells = [];
        for (const cell of elements(row,'c')) {
          const type = cell.getAttribute('t');
          const value = elements(cell,'v')[0]?.textContent ?? '';
          let text = type==='s' ? shared[Number(value)] : type==='inlineStr' ? elements(cell,'t').map(t=>t.textContent).join('') : value;
          text = String(text ?? '');
          const formula = elements(cell,'f').length > 0;
          cells[columnIndex(cell.getAttribute('r') || '')] = {text, type, formula};
        }
        return {number:Number(row.getAttribute('r')), cells};
      });
    }
    return {sheets, rows};
  }
  function identifier(cell, row, label, issues) {
    if (!cell) return '';
    if (cell.type==='e') { issues.push(`Row ${row}: ${label} contains an Excel error.`); return ''; }
    if (cell.formula) { issues.push(`Row ${row}: ${label} is a formula. Paste its value in Excel before converting.`); return ''; }
    const value = cell.text.trim();
    if (!value) return '';
    if (!cell.type || cell.type==='n') {
      // Excel numeric identifiers beyond 15 digits may already be rounded at source.
      if (!/^\d+(?:\.0+)?$/.test(value) || value.split('.')[0].length>15) {
        issues.push(`Row ${row}: ${label} is a numeric identifier that may have lost precision. Re-export it as text.`);
      }
      return value.replace(/\.0+$/, '');
    }
    return value;
  }
  function convert(rows, selection) {
    const {marketplace,brand,country} = selection;
    const channel = CHANNELS[marketplace]?.[brand];
    if (!channel || !['MY','SG'].includes(country)) throw new Error('Select a supported marketplace, country, and brand.');
    const required = marketplace==='Shopee' ? ['Product ID','Variation ID','SKU'] : marketplace==='Lazada' ? ['Product ID','SellerSKU'] : ['CatalogProductId','SellerSku'];
    const header = rows.find(row => required.every(name => row.cells.some(cell=>normalize(cell?.text)===normalize(name))));
    if (!header) throw new Error(`Required columns not found: ${required.join(', ')}. Select the correct marketplace and worksheet.`);
    const index = {};
    required.concat('Brand','Product Name','Parent SKU').forEach(name => {
      const matches = header.cells.map((c,i)=>normalize(c?.text)===normalize(name)?i:-1).filter(i=>i>=0);
      if (matches.length>1) throw new Error('Duplicate column header: '+name);
      index[name] = matches[0] ?? -1;
    });
    let start = rows.indexOf(header)+1, skippedInstructions = 0;
    // Lazada bulk-edit exports have a metadata block: Optional/Mandatory, descriptions, rules.
    if (marketplace==='Lazada' && rows[start]?.cells.some(c=>['optional','mandatory'].includes(c?.text.trim().toLowerCase()))) {
      start += 3; skippedInstructions = 3;
    }
    // Shopee adds leading instructions in price/stock columns. Skip only rows
    // with no product identity, before the first product record.
    if (marketplace==='Shopee') {
      const identityFields=['Product ID','Variation ID','SKU','Product Name','Parent SKU'];
      while (start<rows.length && identityFields.every(name=>
        index[name]===undefined || index[name]<0 || !rows[start].cells[index[name]]?.text.trim()
      )) {
        if (rows[start].cells.some(cell=>cell?.text.trim())) skippedInstructions++;
        start++;
      }
    }
    const records=[], issues=[];
    let blankSkus=0, emptyRows=0, duplicateMappings=0;
    const seen = new Set();
    for (const row of rows.slice(start)) {
      if (!row.cells.some(c=>c?.text.trim())) { emptyRows++; continue; }
      const idField = required[0], skuField = required[required.length-1];
      const product = identifier(row.cells[index[idField]],row.number,idField,issues);
      let sku = identifier(row.cells[index[skuField]],row.number,skuField,issues);
      if (!product) issues.push(`Row ${row.number}: missing ${idField}.`);
      if (!/^\d+$/.test(product) && product) issues.push(`Row ${row.number}: ${idField} must contain digits only.`);
      let productId = product;
      if (marketplace==='Shopee') {
        const variation = identifier(row.cells[index['Variation ID']],row.number,'Variation ID',issues);
        if (!variation) issues.push(`Row ${row.number}: missing Variation ID. Enter the platform's value; it cannot be inferred.`);
        else if (!/^\d+$/.test(variation)) issues.push(`Row ${row.number}: Variation ID must contain digits only.`);
        productId += '~'+variation;
      }
      if (marketplace==='Zalora' && index.Brand>=0) {
        const sourceBrand=row.cells[index.Brand]?.text.trim();
        if (sourceBrand && normalize(sourceBrand)!==normalize(brand)) issues.push(`Row ${row.number}: brand ${sourceBrand} does not match selected ${brand}. Use a single-brand export.`);
      }
      if (!sku) { sku='0'; blankSkus++; }
      const key=JSON.stringify([sku,productId]);
      if (seen.has(key)) duplicateMappings++;
      seen.add(key);
      records.push([channel,sku,sku,productId,'','','','','Active','Create']);
    }
    if (issues.length) {
      const error=new Error(`${issues.length} data issue(s). Fix these before export:\n${issues.slice(0,12).join('\n')}${issues.length>12?'\nAdditional issues omitted.':''}`);
      error.issues=issues; throw error;
    }
    if (!records.length) throw new Error('No product records found below the headers.');
    return {records,blankSkus,duplicateMappings,emptyRows,skippedInstructions,selection:{...selection},batches:Math.ceil(records.length/LIMIT)};
  }
  function batches(result) {
    const {marketplace,country,brand}=result.selection;
    const files=[];
    for (let offset=0;offset<result.records.length;offset+=LIMIT) files.push({
      name:`SkuChannel_${marketplace}_${country}_${brand}_Part${String(files.length+1).padStart(2,'0')}.xls`,
      records:result.records.slice(offset,offset+LIMIT)
    });
    return files;
  }
  async function exportBatch(template, records) {
    if (!records.length || records.length>LIMIT) throw new Error('An export must contain 1–5,198 product records.');
    const zip=await JSZip.loadAsync(template);
    const path='xl/worksheets/sheet1.xml';
    const source=await zip.file(path).async('string');
    const data=source.match(/<sheetData[^>]*>([\s\S]*?)<\/sheetData>/);
    const headers=data?.[1].match(/<row\b[^>]*r="[12]"[^>]*>[\s\S]*?<\/row>/g);
    if (!headers || headers.length!==2) throw new Error('The bundled template is invalid.');
    // Match the original template's shared-string storage for eRetail readers.
    // Append entries so all existing header/support-sheet indexes remain valid.
    const sharedPath='xl/sharedStrings.xml';
    const sharedFile=zip.file(sharedPath);
    if (!sharedFile) throw new Error('The bundled template is missing its shared string table.');
    const sharedSource=await sharedFile.async('string');
    const existingCount=(sharedSource.match(/<si\b/g)||[]).length;
    const added=[], indexes=new Map();
    function sharedIndex(value) {
      const text=String(value);
      if (!indexes.has(text)) {
        indexes.set(text,existingCount+added.length);
        added.push('<si><t xml:space="preserve">'+escapeXml(text)+'</t></si>');
      }
      return indexes.get(text);
    }
    const output=records.map((record,i)=>{
      if (record.length!==10) throw new Error('Invalid output column count.');
      const r=i+3;
      return `<row r="${r}" spans="1:10">`+record.map((value,j)=>{
        const ref=String.fromCharCode(65+j)+r;
        // Text identifiers retain leading zeros; formula-like SKUs remain plain text.
        return value==='' ? `<c r="${ref}"/>` : `<c r="${ref}" t="s"><v>${sharedIndex(value)}</v></c>`;
      }).join('')+'</row>';
    }).join('');
    zip.file(path,source.replace(data[0],'<sheetData>'+headers.join('')+output+'</sheetData>').replace(/<dimension ref="[^"]*"\s*\/>/,`<dimension ref="A1:J${records.length+2}"/>`));
    let references=0;
    for (const entry of Object.values(zip.files)) {
      if (!entry.dir && /^xl\/worksheets\/[^/]+\.xml$/.test(entry.name)) {
        references+=((await entry.async('string')).match(/<c\b[^>]*\bt="s"/g)||[]).length;
      }
    }
    zip.file(sharedPath,sharedSource.replace(/<sst\b[^>]*>/,tag=>tag
      .replace(/\s+count="[^"]*"/,'').replace(/\s+uniqueCount="[^"]*"/,'')
      .replace(/>$/,` count="${references}" uniqueCount="${existingCount+added.length}">`))
      .replace('</sst>',added.join('')+'</sst>'));
    return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  }
  root.SkuConverter={CHANNELS,HEADERS,LIMIT,openWorkbook,convert,batches,exportBatch};
})(globalThis);
