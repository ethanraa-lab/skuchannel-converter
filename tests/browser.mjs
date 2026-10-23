import assert from 'node:assert/strict';
import {createServer} from 'node:http';
import {readFile,stat,mkdir} from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {createRequire} from 'node:module';
const require=createRequire(import.meta.url);
const {chromium}=process.env.CODEX_PRIMARY_RUNTIME_NODE_MODULES ? require(path.join(process.env.CODEX_PRIMARY_RUNTIME_NODE_MODULES,'playwright')) : require('playwright');
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const JSZip=require('../vendor/jszip.min.js');
const server=createServer(async(req,res)=>{
  try{
    const filename=path.join(root,decodeURIComponent(new URL(req.url,'http://localhost').pathname==='/'?'/index.html':new URL(req.url,'http://localhost').pathname));
    if(!filename.startsWith(root+path.sep))throw new Error('invalid path');
    await stat(filename);
    const type=filename.endsWith('.js')?'text/javascript':filename.endsWith('.css')?'text/css':filename.endsWith('.html')?'text/html':'application/octet-stream';
    res.writeHead(200,{'Content-Type':type});res.end(await readFile(filename));
  }catch{res.writeHead(404);res.end('Not found');}
});
await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
const browser=await chromium.launch({headless:true});
const page=await browser.newPage({acceptDownloads:true});
let checks=0;const errors=[];page.on('pageerror',error=>errors.push(error.message));
const check=(condition,message)=>{assert.ok(condition,message);checks++;};
const encodeXml=s=>String(s).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;');
async function fixture(rows){
  const zip=new JSZip();const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  zip.file('xl/workbook.xml',`<workbook xmlns="${ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Products" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml',`<worksheet xmlns="${ns}"><sheetData>`+rows.map((row,i)=>`<row r="${i+1}">`+row.map((v,j)=>`<c r="${String.fromCharCode(65+j)}${i+1}" t="inlineStr"><is><t>${encodeXml(v)}</t></is></c>`).join('')+'</row>').join('')+'</sheetData></worksheet>');
  return zip.generateAsync({type:'nodebuffer'});
}
try{
  await page.goto(`http://127.0.0.1:${server.address().port}`);
  await page.waitForFunction(()=>!!globalThis.SkuConverter);
  const mapping=await page.evaluate(async()=>{
    const C=SkuConverter;
    const rows=(data)=>data.map((r,i)=>({number:i+1,cells:r.map(text=>({text:String(text),type:'s',formula:false}))}));
    const run=(marketplace,brand,data)=>C.convert(rows(data),{marketplace,brand,country:'SG'});
    const sh=run('Shopee','Vans',[['SKU','Variation ID','Product ID'],['00123','00456','00789'],['','9','8']]);
    const la=run('Lazada','Timberland',[['sku.skuId','extra','SellerSKU'],['Optional','','Optional'],['description'],['rules'],['123','','ABC']]);
    const za=run('Zalora','Timberland',[['CatalogProductId','SellerSku','Brand'],['123','ABC','Timberland']]);
    const failures=[];
    for(const [name,fn] of [
      ['missing ID',()=>run('Shopee','Vans',[['SKU','Variation ID','Product ID'],['ABC','','123']])],
      ['wrong columns',()=>run('Shopee','Vans',[['wrong'],['data']])],
      ['wrong brand',()=>run('Zalora','Vans',[['CatalogProductId','SellerSku','Brand'],['123','ABC','Timberland']])],
      ['numeric precision',()=>C.convert([{number:1,cells:['Product ID','Variation ID','SKU'].map(text=>({text,type:'s'}))},{number:2,cells:[{text:'1234567890123456',type:'n'},{text:'2',type:'n'},{text:'ABC',type:'s'}]}],{marketplace:'Shopee',brand:'Vans',country:'MY'})],
      ['formula ID',()=>C.convert([{number:1,cells:['Product ID','Variation ID','SKU'].map(text=>({text,type:'s'}))},{number:2,cells:[{text:'123',type:'n',formula:true},{text:'2',type:'n'},{text:'ABC',type:'s'}]}],{marketplace:'Shopee',brand:'Vans',country:'MY'})]
    ]){try{fn();failures.push(name);}catch{}}
    const large={...sh,records:Array.from({length:12000},(_,i)=>['M02','S'+i,'S'+i,'123~'+i,'','','','','Active','Create'])};
    return {sh,la,za,failures,batches:C.batches(large).map(x=>({name:x.name,count:x.records.length}))};
  });
  check(mapping.sh.records[0][3]==='00789~00456','Shopee leading zeros and concatenation');
  check(mapping.sh.records[1][1]==='0'&&mapping.sh.records[1][2]==='0'&&mapping.sh.blankSkus===1,'blank SKU replacement');
  check(mapping.sh.records[0][0]==='M02'&&mapping.sh.records[0][8]==='Active'&&mapping.sh.records[0][9]==='Create','constants');
  check(mapping.la.records.length===1&&mapping.la.skippedInstructions===3&&mapping.la.records[0][0]==='M05','Lazada instruction rows');
  check(mapping.za.records[0][0]==='M16','Zalora code');
  check(mapping.failures.length===0,'invalid records must block conversion: '+mapping.failures.join(','));
  check(JSON.stringify(mapping.batches.map(x=>x.count))==='[5198,5198,1604]','12000-record split');
  check(mapping.batches[0].name==='SkuChannel_Shopee_SG_Vans_Part01.xls','filename');
  const template=await readFile(path.join(root,'template.xls'));
  const exported=await page.evaluate(async base64=>{
    const source=Uint8Array.from(atob(base64),c=>c.charCodeAt(0));
    const record=['M02','=VALUE<&"','00123','123~456','','','','','Active','Create'];
    const bytes=await SkuConverter.exportBatch(source,[record]);
    return Array.from(bytes);
  },template.toString('base64'));
  const output=await JSZip.loadAsync(Uint8Array.from(exported)), original=await JSZip.loadAsync(template);
  for(const [name,file] of Object.entries(original.files)){
    if(file.dir||(name==='xl/worksheets/sheet1.xml'||name==='xl/sharedStrings.xml'))continue;
    assert.deepEqual(await output.file(name).async('uint8array'),await file.async('uint8array'),'preserved component '+name);
  }
  checks++;
  const outXml=await output.file('xl/worksheets/sheet1.xml').async('string');
  const originalXml=await original.file('xl/worksheets/sheet1.xml').async('string');
  const headerRows=text=>text.match(/<row\b[^>]*r="[12]"[^>]*>[\s\S]*?<\/row>/g);
  check(JSON.stringify(headerRows(outXml))===JSON.stringify(headerRows(originalXml)),'exact headers');
  const sharedXml=await output.file('xl/sharedStrings.xml').async('string');
  check(!outXml.includes('<f>')&&sharedXml.includes('=VALUE&lt;&amp;&quot;'),'formula-like strings escaped as text');
  check(outXml.includes('A1:J3'),'dimension');
  const sampleDir=process.env.SAMPLE_DIR;
  if(sampleDir){
    for(const [market,brand,count,missing] of [['Shopee','Vans',5,1],['Lazada','Vans',6,0],['Zalora','Timberland',12,0]]){
      const bytes=await readFile(path.join(sampleDir,market+'LayoutExample.xlsx'));
      const actual=await page.evaluate(async({base64,market,brand})=>{
        const w=await SkuConverter.openWorkbook(Uint8Array.from(atob(base64),c=>c.charCodeAt(0)));
        return SkuConverter.convert(await w.rows(0),{marketplace:market,country:'MY',brand});
      },{base64:bytes.toString('base64'),market,brand});
      check(actual.records.length===count,market+' sample record count');
      check(actual.blankSkus===missing,market+' sample blanks');
    }
  }
  const bytes=await fixture([['Product ID','Variation ID','SKU'],['123','456','001SKU'],['123','789','']]);
  await page.selectOption('#channel','M02');
  await page.setInputFiles('#file',{name:'synthetic.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:bytes});
  await page.waitForFunction(()=>!document.getElementById('convert').disabled);
  await page.click('#convert');await page.waitForSelector('#results:not([hidden])');
  check(await page.locator('#record-count').textContent()==='2','UI conversion');
  const downloadPromise=page.waitForEvent('download');await page.locator('.batch button').first().click();const download=await downloadPromise;
  check(download.suggestedFilename()==='SkuChannel_Shopee_MY_Vans_Part01.xls','individual download');
  await download.path();
  await page.waitForFunction(()=>!document.getElementById('download-all').disabled);
  const zipPromise=page.waitForEvent('download');await page.click('#download-all');const zipDownload=await zipPromise;
  const zipBytes=await readFile(await zipDownload.path());const zip=await JSZip.loadAsync(zipBytes);
  check(Object.keys(zip.files).length===1&&!!zip.file(download.suggestedFilename()),'ZIP contains complete batch');
  await page.waitForFunction(()=>!document.getElementById('country').disabled);
  await page.check('input[name="country"][value="SG"]');check(await page.locator('#results').isHidden(),'selection change invalidates old results');
  await page.click('#convert');await page.waitForSelector('#results:not([hidden])');
  await mkdir(path.join(root,'test-results'),{recursive:true});
  await page.screenshot({path:path.join(root,'test-results','desktop.png'),fullPage:true});
  await page.setViewportSize({width:390,height:844});await page.screenshot({path:path.join(root,'test-results','mobile.png'),fullPage:true});
  check(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),'mobile has no page overflow');
  check(errors.length===0,'no browser errors: '+errors.join(';'));
  console.log(`PASS: ${checks} checks, including browser downloads and exact template preservation.`);
}finally{await browser.close();server.close();}
