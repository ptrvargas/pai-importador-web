/* Importador OOXML: edición quirúrgica de celdas, sin reserializar libros. */
export const OPTIONS = ['Absolutamente Falso','Ligeramente Cierto','Moderadamente Cierto','Completamente Cierto'];
export const normalize = value => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim().replace(/\s+/g,' ');
const fail = message => { throw new Error(message); };
export function decode(s='') { return s.replace(/&#x([\da-f]+);|&#(\d+);|&(amp|lt|gt|quot|apos);/gi,(_,h,d,k)=>h?String.fromCodePoint(parseInt(h,16)):d?String.fromCodePoint(+d):({amp:'&',lt:'<',gt:'>',quot:'"',apos:"'"})[k.toLowerCase()]); }
const escape = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
function attrs(s) { return Object.fromEntries([...s.matchAll(/([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)].map(m=>[m[1],decode(m[2]??m[3])])); }
function tags(xml,name) { const re = new RegExp(`<((?:[\\w.-]+:)?${name})\\b([^>]*?)(?:\\/\\s*>|>([\\s\\S]*?)<\\/\\1\\s*>)`,'g'); return [...xml.matchAll(re)].map(m=>({raw:m[0],tag:m[1],a:attrs(m[2]),body:m[3]??''})); }
function text(xml) { return tags(xml,'t').map(t=>decode(t.body)).join(''); }
export function cells(xml,strings=[]) { const out=new Map(); for(const c of tags(xml,'c')) { const v=tags(c.body,'v')[0]?.body??''; out.set(c.a.r,{...c,value:c.a.t==='s'?(strings[+v]??fail('El archivo contiene una referencia de texto inválida.')):c.a.t==='inlineStr'?text(c.body):decode(v),formula:tags(c.body,'f')[0]?.raw??null}); } return out; }
function pathTarget(target) { const parts=('xl/'+target).split('/');const out=[];if(target.startsWith('/'))return target.slice(1);for(const p of parts){if(p==='..')out.pop();else if(p!=='.')out.push(p);}return out.join('/'); }
export async function openWorkbook(data,JSZip) {
 if(data.byteLength>30*1024*1024)fail('El archivo supera el límite de 30 MB.');
 let zip;try{zip=await JSZip.loadAsync(data);}catch{fail('No se pudo abrir el Excel. Comprueba que no esté cifrado, dañado o protegido con contraseña de apertura.');}
 let size=0;for(const e of Object.values(zip.files)){size+=e._data?.uncompressedSize??0;}if(size>120*1024*1024)fail('El contenido descomprimido supera el límite de seguridad de 120 MB.');
 try{zip=await JSZip.loadAsync(data,{checkCRC32:true});}catch{fail('El archivo está dañado. Vuelve a descargarlo desde su origen.');}
 if(Object.keys(zip.files).some(p=>p.startsWith('_xmlsignatures/')))fail('Este libro tiene una firma digital del documento y no puede modificarse sin invalidarla.');
 const read=async p=>{const f=zip.file(p);if(!f)fail('El archivo no tiene la estructura de Excel esperada.');return f.async('string');};
 const workbook=await read('xl/workbook.xml'), rels=tags(await read('xl/_rels/workbook.xml.rels'),'Relationship');
 const strings=zip.file('xl/sharedStrings.xml')?tags(await read('xl/sharedStrings.xml'),'si').map(s=>text(s.body)):[];
 const sheets=[];for(const s of tags(workbook,'sheet')){const rel=rels.find(r=>r.a.Id===s.a['r:id']);if(!rel||rel.a.TargetMode==='External')fail('La estructura de hojas no es compatible.');const path=pathTarget(rel.a.Target),xml=await read(path);sheets.push({name:s.a.name,path,xml,cells:cells(xml,strings)});}
 return {zip,workbook,strings,sheets};
}
export function inspectCorrector(book) {
 if(!book.zip.file('xl/vbaProject.bin'))fail('Selecciona el corrector maestro con sus macros VBA (.xlsm).');
 const candidates=[];
 for(const sheet of book.sheets){const c=sheet.cells;const cols=['F','G','H','I'];if(!cols.every((col,i)=>normalize(c.get(col+'1')?.value)===normalize(OPTIONS[i])))continue;
 const mapping=new Map();for(const [ref,cell] of c){if(!/^D\d+$/.test(ref)||!/^\d+$/.test(cell.value))continue;const id=+cell.value,row=+ref.slice(1);if(id<1||id>344)continue;if(mapping.has(id))fail('El corrector contiene números de pregunta duplicados.');if(!c.get('E'+row)?.value||!c.get('J'+row)?.formula)fail('La estructura de preguntas del corrector no es compatible.');for(const col of cols){if(!c.has(col+row)||c.get(col+row).formula)fail('Las celdas de respuesta del corrector no son compatibles.');}mapping.set(id,row);}
 if(mapping.size===344)candidates.push({sheet,mapping});}
 if(candidates.length!==1)fail('No se encontró una hoja única con las 344 preguntas y las cuatro columnas del corrector PAI.');
 const ids=book.sheets.filter(s=>normalize(s.cells.get('D5')?.value)==='nombres y apellidos'&&normalize(s.cells.get('D6')?.value)==='edad:'&&normalize(s.cells.get('F6')?.value)==='sexo:');
 if(ids.length!==1)fail('No se pudo reconocer la ficha de identificación del corrector.');
 for(const ref of ['F5','I5','E6','G6','I6','F7'])if(!ids[0].cells.has(ref)||ids[0].cells.get(ref).formula)fail('La ficha de identificación tiene una estructura no compatible.');
 return {...candidates[0],identification:ids[0],book};
}
function metadataKey(header){ const h=normalize(header).replace(/[:*]/g,'').trim();if(['nombre completo','nombres y apellidos','nombre y apellidos'].includes(h))return 'name';if(['genero','sexo'].includes(h))return 'sex';if(h==='edad')return 'age';if(['grado de escolaridad','nivel de instruccion','escolaridad'].includes(h))return 'education';if(['timestamp','marca temporal','fecha','fecha de evaluacion'].includes(h))return 'date';return null; }
export function inspectResponses(book){
 const parts=[[],[]];let key=0;
 for(const sheet of book.sheets){const rows=new Map();for(const [ref,c] of sheet.cells){const m=ref.match(/^([A-Z]+)(\d+)$/);if(!m)continue;const r=+m[2];if(!rows.has(r))rows.set(r,new Map());rows.get(r).set(m[1],c);}
 const header=[...rows.entries()].sort((a,b)=>a[0]-b[0]).find(([,r])=>[...r.values()].filter(c=>/^\s*\d+[.)]\s*/.test(c.value)).length>=170);if(!header)continue;
 const [hr,h]=header,questions=new Map(),meta=new Map();for(const [col,c] of h){const m=c.value.match(/^\s*(\d+)[.)]\s*/);if(m){const id=+m[1];if(questions.has(id))fail('El archivo de respuestas contiene preguntas duplicadas.');questions.set(id,col);}else{const k=metadataKey(c.value);if(k){if(meta.has(k))fail('Hay columnas de identificación duplicadas.');meta.set(k,col);}}}
 const ids=[...questions.keys()].sort((a,b)=>a-b);const part=ids[0]===1?0:ids[0]===173?1:-1;if(part<0||ids.length!==172||!ids.every((id,i)=>id===part*172+i+1))fail('Cada parte debe contener exactamente sus 172 preguntas, sin saltos ni duplicados.');
 if(!meta.has('name'))fail('Falta la columna «Nombre completo» en una de las partes.');
 for(const [row,values] of rows){if(row<=hr||![...values.values()].some(c=>String(c.value).trim()))continue;const data={};for(const [k,col] of meta)data[k]=values.get(col)?.value?.trim()??'';const answers={};for(const [id,col] of questions){const cell=values.get(col);answers[id]=cell?.value??'';if(cell?.formula)fail('Las respuestas deben ser valores del formulario, no fórmulas.');}parts[part].push({key:String(key++),part,row,sheet:sheet.name,data,answers});}
 }
 if(!parts[0].length||!parts[1].length)fail('No se encontraron evaluaciones en ambas partes (1–172 y 173–344).');return parts;
}
export function autoPair(first,parts){const name=normalize(first.data.name);if(!name)return null;const a=parts[0].filter(x=>normalize(x.data.name)===name),b=parts[1].filter(x=>normalize(x.data.name)===name);return a.length===1&&b.length===1?b[0]:null;}
export function requiresConfirmation(first,second,parts){return !normalize(first.data.name)||!normalize(second.data.name)||normalize(first.data.name)!==normalize(second.data.name)||autoPair(first,parts)?.key!==second.key;}
export function validatePair(first,second,parts,confirmed=false){
 if(!first||!second||first.part!==0||second.part!==1)fail('Selecciona una evaluación de cada parte.');
 if(!first.data.name.trim()||!second.data.name.trim())fail('Ambas partes deben tener un nombre. Corrige el archivo de origen.');
 if(requiresConfirmation(first,second,parts)&&!confirmed)fail('Los nombres son diferentes o hay evaluaciones repetidas. Revisa ambas partes y confirma expresamente que corresponden a la misma evaluación.');
 for(const k of ['sex','age','education'])if(first.data[k]&&second.data[k]&&normalize(first.data[k])!==normalize(second.data[k]))fail('Las partes tienen datos de identificación contradictorios. Revisa el archivo de origen.');
 const answers={...first.answers,...second.answers},errors=[],codes=[];for(let i=1;i<=344;i++){const value=answers[i],code=OPTIONS.findIndex(x=>normalize(x)===normalize(value));if(code<0)errors.push(i);codes.push(code);}
 if(errors.length)fail(`Hay ${errors.length} respuestas vacías o no permitidas. Revisa las preguntas: ${errors.slice(0,20).join(', ')}${errors.length>20?'…':''}.`);
 const data={...second.data,...Object.fromEntries(Object.entries(first.data).filter(([,v])=>v!==''))};const age=Number(data.age);if(data.age&&(!Number.isInteger(age)||age<1||age>120))fail('La edad no es válida. Revisa el formulario.');
 const sex={masculino:'M',femenino:'F',m:'M',f:'F'}[normalize(data.sex)];if(data.sex&&!sex)fail('El género no es compatible con la ficha original (M/F).');
 return {codes,data:{...data,sex:sex??'',age:data.age?String(age):''}};
}
function replaceCells(xml,changes){const found=new Set();const out=xml.replace(/<((?:[\w.-]+:)?c)\b[^>]*?(?:\/\s*>|>[\s\S]*?<\/\1\s*>)/g,raw=>{const open=raw.match(/^<[^>]+>/)[0],a=attrs(open);if(!changes.has(a.r))return raw;const value=changes.get(a.r);found.add(a.r);if(tags(raw,'f').length)fail('Se intentó modificar una fórmula. Operación cancelada.');const tag=raw.match(/^<([^\s/>]+)/)[1],prefix=tag.includes(':')?tag.split(':')[0]+':':'';let cleaned=open.replace(/\s+t\s*=\s*(?:"[^"]*"|'[^']*')/g,'').replace(/\s*\/?>$/,'');const rest=raw.slice(open.length).replace(/<\/(?:[\w.-]+:)?c\s*>$/,'').replace(/<((?:[\w.-]+:)?(?:v|is))\b[^>]*?(?:\/\s*>|>[\s\S]*?<\/\1\s*>)/g,'');return typeof value==='number'?`${cleaned}><${prefix}v>${value}</${prefix}v>${rest}</${tag}>`:value===''?`${cleaned}>${rest}</${tag}>`:`${cleaned} t="inlineStr"><${prefix}is><${prefix}t xml:space="preserve">${escape(value)}</${prefix}t></${prefix}is>${rest}</${tag}>`;});if(found.size!==changes.size)fail('Falta una celda necesaria en el corrector.');return out;}
const sameBytes=(a,b)=>a.length===b.length&&a.every((v,i)=>v===b[i]);
function formulas(xml){return tags(xml,'f').map(x=>x.raw).join('\n');}
export async function generate(corrector,validated,JSZip){
 const {book,sheet,mapping,identification}=corrector;const changes=new Map();
 for(let i=1;i<=344;i++)for(let c=0;c<4;c++)changes.set('FGHI'[c]+mapping.get(i),validated.codes[i-1]===c?'X':'');
 const questions=replaceCells(sheet.xml,changes);
 let ident=identification.xml;const data=validated.data;
 let date=data.date??'';if(date&&!/^\d+(?:\.\d+)?$/.test(date)){const match=date.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?$/);if(match)date=String(Math.floor((Date.UTC(+match[1],+match[2]-1,+match[3])-Date.UTC(1899,11,30))/86400000));}
 const fields={F5:data.name,E6:data.age?Number(data.age):'',G6:data.sex,I6:'',F7:data.education??'',I5:date&&/^\d+(?:\.\d+)?$/.test(date)?Math.floor(Number(date)):date};
 ident=replaceCells(ident,new Map(Object.entries(fields)));
 let workbook=book.workbook;if(/<(?:[\w.-]+:)?calcPr\b/.test(workbook))workbook=workbook.replace(/<((?:[\w.-]+:)?calcPr)\b([^>]*?)(\/?>)/,(_,tag,a,end)=>`<${tag}${a.replace(/\s+(?:fullCalcOnLoad|forceFullCalc)="[^"]*"/g,'')} fullCalcOnLoad="1" forceFullCalc="1"${end}`);else fail('El corrector no tiene configuración de cálculo compatible.');
 const zip=await JSZip.loadAsync(await book.zip.generateAsync({type:'uint8array'}));zip.file(sheet.path,questions,{createFolders:false});zip.file(identification.path,ident,{createFolders:false});zip.file('xl/workbook.xml',workbook,{createFolders:false});
 const output=await zip.generateAsync({type:'uint8array',compression:'DEFLATE',compressionOptions:{level:6}});
 const checked=await openWorkbook(output,JSZip);const changed=new Set([sheet.path,identification.path,'xl/workbook.xml']);
 if(Object.keys(book.zip.files).sort().join('|')!==Object.keys(checked.zip.files).sort().join('|'))fail('La comprobación de componentes del archivo no fue satisfactoria.');
 for(const [path,entry] of Object.entries(book.zip.files)){if(entry.dir)continue;if(!changed.has(path)){if(!sameBytes(await entry.async('uint8array'),await checked.zip.file(path).async('uint8array')))fail('No se pudo garantizar la integridad de los componentes originales.');}else if(formulas(await entry.async('string'))!==formulas(await checked.zip.file(path).async('string')))fail('Se detectó un cambio de fórmulas. Descarga cancelada.');}
 const result=inspectCorrector(checked);for(let i=1;i<=344;i++){const values=[...'FGHI'].map(c=>result.sheet.cells.get(c+result.mapping.get(i))?.value);if(values.filter(x=>x==='X').length!==1||values[validated.codes[i-1]]!=='X')fail('La comprobación de las 344 marcas no fue satisfactoria.');}
 return output;
}
