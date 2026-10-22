'use strict';
const $=id=>document.getElementById(id);
let workbook=null,result=null,templatePromise=null,busy=false,loadVersion=0;
const converter=globalThis.SkuConverter;
function message(text,error=false){$('message').textContent=text;$('message').className=error?'error':'';}
function reset(){result=null;$('results').hidden=true;$('empty').hidden=false;$('download-all').disabled=true;message('');}
function updateCode(){ $('channel-code').textContent=converter.CHANNELS[$('marketplace').value][$('brand').value]; }
function updateBrands(){
  const previous=$('brand').value;
  $('brand').replaceChildren(...Object.keys(converter.CHANNELS[$('marketplace').value]).map(brand=>new Option(brand,brand)));
  if (converter.CHANNELS[$('marketplace').value][previous]) $('brand').value=previous;
  updateCode();reset();
}
function setBusy(value){
  busy=value;
  for(const id of ['marketplace','country','brand','file','sheet']) $(id).disabled=value;
  $('convert').disabled=value||!workbook;
  $('download-all').disabled=value||!result;
  document.querySelectorAll('.batch button').forEach(button=>button.disabled=value);
}
$('marketplace').addEventListener('change',updateBrands);
for(const id of ['country','brand','sheet']) $(id).addEventListener('change',()=>{updateCode();reset();});
$('file').addEventListener('change',async()=>{
  const version=++loadVersion;reset();workbook=null;$('sheet-container').hidden=true;$('convert').disabled=true;
  const file=$('file').files[0];$('file-name').textContent=file?.name||'No file selected';if(!file)return;
  setBusy(true);message('Reading workbook…');
  try {
    if(file.size>30*1024*1024)throw new Error('File exceeds 30 MB. Export a smaller file.');
    workbook=await converter.openWorkbook(await file.arrayBuffer());
    if(version!==loadVersion)return;
    $('sheet').replaceChildren(...workbook.sheets.map((sheet,i)=>new Option(sheet.name,String(i))));
    $('sheet-container').hidden=workbook.sheets.length<2;message('Workbook ready. Check the channel selection, then convert.');
  }catch(error){workbook=null;message(error.message,true);}finally{if(version===loadVersion)setBusy(false);}
});
$('convert').addEventListener('click',async()=>{
  if(busy||!workbook)return;reset();setBusy(true);message('Converting records…');
  try{
    await new Promise(resolve=>setTimeout(resolve,0));
    const selection={marketplace:$('marketplace').value,country:$('country').value,brand:$('brand').value};
    result=converter.convert(await workbook.rows(Number($('sheet').value)),selection);render();message('Conversion complete. Review the preview before downloading.');
  }catch(error){message(error.message,true);}finally{setBusy(false);}
});
function render(){
  $('results').hidden=false;$('empty').hidden=true;
  $('record-count').textContent=result.records.length.toLocaleString();$('batch-count').textContent=result.batches;$('blank-count').textContent=result.blankSkus;
  $('summary').textContent=`${result.selection.marketplace} ${result.selection.country} · ${result.selection.brand} · ${result.skippedInstructions} instruction rows skipped · ${result.emptyRows} empty rows skipped`;
  $('warnings').replaceChildren();
  const warnings=[];
  if(result.blankSkus)warnings.push(`${result.blankSkus} blank SKU(s) were replaced with 0 in both SKU columns. Check these records before importing into eRetail.`);
  if(result.duplicateMappings)warnings.push(`${result.duplicateMappings} repeated SKU/product mapping(s) retained. No records were deduplicated.`);
  for(const text of warnings){const p=document.createElement('p');p.className='warning';p.textContent=text;$('warnings').append(p);}
  $('preview').replaceChildren(...result.records.slice(0,20).map(record=>{
    const tr=document.createElement('tr');for(const i of [0,1,2,3,8,9]){const td=document.createElement('td');td.textContent=record[i];tr.append(td);}return tr;
  }));
  $('batches').replaceChildren(...converter.batches(result).map(batch=>{
    const div=document.createElement('div');div.className='batch';const span=document.createElement('span');span.textContent=batch.name;
    const small=document.createElement('small');small.textContent=`${batch.records.length.toLocaleString()} records + 2 header rows`;span.append(small);
    const button=document.createElement('button');button.textContent='Download';button.setAttribute('aria-label','Download '+batch.name);
    button.addEventListener('click',()=>downloadBatch(batch));div.append(span,button);return div;
  }));
}
async function template(){
  if(!templatePromise)templatePromise=fetch('template.xls').then(response=>{if(!response.ok)throw new Error('Could not load the eRetail template.');return response.arrayBuffer();}).catch(error=>{templatePromise=null;throw error;});
  return templatePromise;
}
function save(bytes,name,type){
  const url=URL.createObjectURL(new Blob([bytes],{type}));const a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),10000);
}
async function downloadBatch(batch){
  if(busy)return;setBusy(true);message('Preparing '+batch.name+'…');
  try{save(await converter.exportBatch(await template(),batch.records),batch.name,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');message('File prepared. Check your browser downloads.');}
  catch(error){message(error.message,true);}finally{setBusy(false);}
}
$('download-all').addEventListener('click',async()=>{
  if(busy||!result)return;setBusy(true);
  try{
    const zip=new JSZip(),source=await template();const files=converter.batches(result);
    for(let i=0;i<files.length;i++){message(`Preparing file ${i+1} of ${files.length}…`);zip.file(files[i].name,await converter.exportBatch(source,files[i].records));}
    const {marketplace,country,brand}=result.selection;
    save(await zip.generateAsync({type:'uint8array',compression:'STORE'}),`SkuChannel_${marketplace}_${country}_${brand}.zip`,'application/zip');message('ZIP prepared. Check your browser downloads.');
  }catch(error){message(error.message,true);}finally{setBusy(false);}
});
updateBrands();
