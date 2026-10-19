import assert from 'node:assert/strict';
import {readFile,writeFile,mkdir} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import '../vendor/jszip.min.js';
import {openWorkbook,inspectCorrector,inspectResponses,normalize,autoPair,requiresConfirmation,validatePair,generate,OPTIONS} from '../core.js';
const JSZip=globalThis.JSZip;
const args=process.argv.slice(2),source=args[0],responses=args[1],destination=args[2];
if(!source||!responses)throw new Error('Uso: node tests/core.test.mjs /ruta/corrector.xlsm /ruta/respuestas.xlsx [/ruta/privada/salida.xlsm]');
const masterBytes=await readFile(source),responseBytes=await readFile(responses),hash=b=>createHash('sha256').update(b).digest('hex');
const book=await openWorkbook(masterBytes,JSZip),corrector=inspectCorrector(book),rbook=await openWorkbook(responseBytes,JSZip),parts=inspectResponses(rbook);
const first=parts[0][0],second=parts[1][0];let count=0;
function test(name,fn){fn();count++;console.log('OK',name);}
async function asyncTest(name,fn){await fn();count++;console.log('OK',name);}
test('344 preguntas y 12 hojas del corrector',()=>{assert.equal(corrector.mapping.size,344);assert.equal(book.sheets.length,12);});
test('Dos partes completas detectadas por estructura',()=>{assert.equal(Object.keys(first.answers).length,172);assert.equal(Object.keys(second.answers).length,172);});
test('Los nombres reales diferentes exigen confirmación',()=>{assert.notEqual(normalize(first.data.name),normalize(second.data.name));assert.equal(autoPair(first,parts),null);assert.equal(requiresConfirmation(first,second,parts),true);assert.throws(()=>validatePair(first,second,parts),/confirma/);});
const validated=validatePair(first,second,parts,true);
test('344 opciones válidas tras confirmación explícita de prueba',()=>assert.equal(validated.codes.length,344));
test('Mayúsculas, tildes y espacios no cambian la identidad',()=>{const a={...first,data:{...first.data,name:'  MARÍA   pérez '}},b={...second,data:{...second.data,name:'maria PEREZ'}};assert.equal(autoPair(a,[[a],[b]]),b);assert.equal(requiresConfirmation(a,b,[[a],[b]]),false);});
test('Nombres repetidos no se asocian automáticamente',()=>{const b={...second,data:{...second.data,name:first.data.name}},duplicate={...first,key:'duplicada'};assert.equal(autoPair(first,[[first,duplicate],[b]]),null);assert.throws(()=>validatePair(first,b,[[first,duplicate],[b]]),/confirma/);});
test('Sin nombre no se genera ni con confirmación',()=>assert.throws(()=>validatePair({...first,data:{...first.data,name:''}},second,parts,true),/nombre/));
test('Respuestas vacías y opciones desconocidas bloqueadas',()=>{for(const value of ['', 'A veces', '0'])assert.throws(()=>validatePair({...first,answers:{...first.answers,1:value}},second,parts,true),/preguntas: 1/);});
test('Datos contradictorios bloqueados',()=>assert.throws(()=>validatePair(first,{...second,data:{...second.data,age:'85'}},parts,true),/contradictorios/));
await asyncTest('Archivos intercambiados bloqueados',async()=>{assert.throws(()=>inspectCorrector(rbook),/macros/);assert.throws(()=>inspectResponses(book),/ambas partes/);});
await asyncTest('Archivo dañado bloqueado',async()=>assert.rejects(openWorkbook(new Uint8Array([1,2,3]),JSZip),/No se pudo abrir/));
await asyncTest('Archivo grande bloqueado antes de descomprimir',async()=>assert.rejects(openWorkbook(new Uint8Array(30*1024*1024+1),JSZip),/30 MB/));
const output=await generate(corrector,validated,JSZip),out=await openWorkbook(output,JSZip),outCorrector=inspectCorrector(out);
test('344 marcas correctas, ninguna duplicada',()=>{for(let i=1;i<=344;i++){const row=outCorrector.mapping.get(i);const values=[...'FGHI'].map(c=>outCorrector.sheet.cells.get(c+row).value);assert.equal(values.filter(x=>x==='X').length,1);assert.equal(values[validated.codes[i-1]],'X');}});
await asyncTest('VBA idéntico byte a byte',async()=>assert.equal(hash(await book.zip.file('xl/vbaProject.bin').async('uint8array')),hash(await out.zip.file('xl/vbaProject.bin').async('uint8array'))));
await asyncTest('Solo cambian cuestionario, identificación y solicitud de recálculo',async()=>{let changed=[];for(const [p,e] of Object.entries(book.zip.files))if(!e.dir&&hash(await e.async('uint8array'))!==hash(await out.zip.file(p).async('uint8array')))changed.push(p);assert.deepEqual(changed.sort(),[corrector.sheet.path,corrector.identification.path,'xl/workbook.xml'].sort());});
test('Fórmulas, hojas, gráficos y baremos conservados',()=>{assert.deepEqual(book.sheets.map(x=>x.name),out.sheets.map(x=>x.name));for(let i=0;i<book.sheets.length;i++){assert.deepEqual([...book.sheets[i].cells].filter(([,c])=>c.formula).map(([ref,c])=>[ref,c.formula]),[...out.sheets[i].cells].filter(([,c])=>c.formula).map(([ref,c])=>[ref,c.formula]));}});
test('Identificación numérica y datos de ejemplo limpiados',()=>{const c=outCorrector.identification.cells;assert.equal(c.get('F5').value,first.data.name);assert.equal(c.get('E6').value,'48');assert.equal(c.get('E6').a.t,undefined);assert.equal(c.get('G6').value,'M');assert.equal(c.get('F7').value,'Posgrado');assert.equal(c.get('I6').value,'');assert.equal(c.get('I5').value,'46301');});
test('El original conserva sus marcas de ejemplo',()=>{assert.notEqual(corrector.sheet.cells.get('I2').value,outCorrector.sheet.cells.get('I2').value);});
await asyncTest('Nombres de archivo arbitrarios no afectan el reconocimiento',async()=>{const renamed={name:'un nombre con espacios y tildes.xlsm',bytes:masterBytes};assert.equal(inspectCorrector(await openWorkbook(renamed.bytes,JSZip)).mapping.size,344);const r={name:'otro nombre final.xlsx',bytes:responseBytes};assert.equal(inspectResponses(await openWorkbook(r.bytes,JSZip))[0].length,parts[0].length);});
await asyncTest('Cabecera duplicada o incompleta bloqueada',async()=>{for(const next of ['1. duplicada','999. fuera de rango']){const zip=await JSZip.loadAsync(responseBytes);const shared=await zip.file('xl/sharedStrings.xml').async('string');zip.file('xl/sharedStrings.xml',shared.replace(/2\. Tengo luchas internas que me causan problemas/,next));const changed=await openWorkbook(await zip.generateAsync({type:'uint8array'}),JSZip);assert.throws(()=>inspectResponses(changed),/duplicadas|172 preguntas/);}});
await asyncTest('Cuatro opciones se insertan en F, G, H e I',async()=>{const codes=Array.from({length:344},(_,i)=>i%4);const roundtrip=inspectCorrector(await openWorkbook(await generate(corrector,{...validated,codes},JSZip),JSZip));for(let i=1;i<=344;i++)assert.equal(roundtrip.sheet.cells.get('FGHI'[codes[i-1]]+roundtrip.mapping.get(i)).value,'X');});
assert.equal(hash(await readFile(source)),hash(masterBytes));assert.equal(hash(await readFile(responses)),hash(responseBytes));
if(destination){await mkdir(new URL('.',`file://${destination}`).pathname,{recursive:true});await writeFile(destination,output);}
console.log(JSON.stringify({tests:count,result:'PASS',marks:344,sheets:book.sheets.length,vbaSha256:hash(await book.zip.file('xl/vbaProject.bin').async('uint8array')),masterSha256:hash(masterBytes),responsesSha256:hash(responseBytes)},null,2));
