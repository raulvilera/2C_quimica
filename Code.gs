/**
 * Backend da atividade Química — 2ª Série C — 3º Bimestre.
 * Planilha: Atividade de Química 2ªC (3ºBimestre)
 * Aba criada/atualizada automaticamente: Respostas 3B — IA
 *
 * Configure GEMINI_API_KEY em Propriedades do projeto do Apps Script.
 */
const SPREADSHEET_ID = '1bDjs0CREAignhZE5HGNMsp5O9iY_DYAY7QJuHtZyKw0';
const RESULTS_SHEET_NAME = 'Respostas 3B — IA';
const GEMINI_MODEL = 'gemini-3-flash-preview';
const OBJECTIVE_KEY = { q1:'B', q2:'A', q3:'A', q4:'A', q5:'A' };
const ESSAY_RUBRIC = {
  q6:'D1/D2 — balancear Fe2O3 + 3 CO -> 2 Fe + 3 CO2; relacionar 1 mol Fe2O3 a 2 mol Fe; calcular 112 g teóricos e rendimento aproximado de 87,5% para 98 g obtidos.',
  q7:'D2 — calcular 12 mol x 75% = 9 mol ativos; comparar com a razão 3:2; identificar o limitante; separar quantidade teórica, pureza e rendimento.',
  q8:'D3 — propor controle de variáveis; comparar inteiro/triturado na mesma temperatura e frio/morno com mesma granulometria; usar tempo/rapidez; explicar colisões efetivas.',
  q9:'D4 — relacionar pH a H+ e pOH a OH-; explicar equilíbrio dinâmico e deslocamento por concentração/temperatura; não afirmar que o sistema para.',
  q10:'D5 — diferenciar mistura, dissolução e precipitação; comparar produto iônico com Kps; explicar nova fase e sugerir tratamento/monitoramento ambiental.'
};

function doGet(){ return json_({ok:true,service:'atividade-quimica-2C-3bimestre'}); }

function doPost(e){
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const data=JSON.parse((e&&e.postData&&e.postData.contents)||'{}'); validatePayload_(data);
    const objectiveScore=scoreObjective_(data.answers);
    const ai=gradeEssays_(data.essays);
    const essayScore=Number(ai.essayScore||0), totalScore=objectiveScore+essayScore, percentage=Math.round(totalScore/15*100);
    const sheet=ensureResultsSheet_();
    const row=[new Date(),data.date,data.series,data.number,data.student,
      data.answers.q1,data.answers.q2,data.answers.q3,data.answers.q4,data.answers.q5,
      data.essays.q6,data.essays.q7,data.essays.q8,data.essays.q9,data.essays.q10,
      objectiveScore,essayScore,totalScore,percentage+'%',JSON.stringify(ai.scores||{}),ai.feedback||''];
    sheet.appendRow(row);
    return json_({ok:true,student:data.student,objectiveScore,essayScore,totalScore,percentage,feedback:ai.feedback||''});
  }catch(err){ return json_({ok:false,error:String(err.message||err)}); }
  finally{ lock.releaseLock(); }
}

function ensureResultsSheet_(){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh=ss.getSheetByName(RESULTS_SHEET_NAME);
  if(!sh) sh=ss.insertSheet(RESULTS_SHEET_NAME);
  if(sh.getLastRow()===0){
    const headers=['Timestamp','Data','Série','Nº','Aluno','Q1','Q2','Q3','Q4','Q5','Discursiva 6','Discursiva 7','Discursiva 8','Discursiva 9','Discursiva 10','Objetivas /5','Discursivas /10','Total /15','Percentual','Pontuação IA por questão','Devolutiva IA'];
    sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#0d5c91').setFontColor('#ffffff');
    sh.setFrozenRows(1); sh.autoResizeColumns(1,headers.length);
  }
  return sh;
}

function validatePayload_(d){
  if(!d.student||!d.date) throw new Error('Nome e data são obrigatórios.');
  if(!d.answers||Object.keys(d.answers).length!==5) throw new Error('As cinco objetivas são obrigatórias.');
  if(!d.essays||Object.keys(d.essays).length!==5) throw new Error('As cinco discursivas são obrigatórias.');
  Object.keys(d.essays).forEach(k=>{if(String(d.essays[k]).trim().length<15)throw new Error('Resposta insuficiente em '+k+'.');});
}
function scoreObjective_(a){return Object.keys(OBJECTIVE_KEY).reduce((n,k)=>n+(String(a[k]||'').toUpperCase()===OBJECTIVE_KEY[k]?1:0),0);}

function gradeEssays_(essays){
  const key=PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if(!key) throw new Error('Configure GEMINI_API_KEY nas propriedades do projeto Apps Script.');
  const prompt=`Você é avaliador pedagógico de Química do Ensino Médio. Corrija cinco respostas discursivas em estilo SARESP/ENEM. Use apenas a rubrica; não invente conteúdo.\n\nPontuação: 0 = ausente/incompatível; 1 = parcial, com lacunas; 2 = aprendizagem essencial demonstrada com conceitos, dados, unidades e justificativa coerente.\n\nRubrica:\n${Object.entries(ESSAY_RUBRIC).map(([k,v])=>k+': '+v).join('\n')}\n\nRespostas:\n${Object.entries(essays).map(([k,v])=>'\n'+k+': '+v).join('\n')}\n\nResponda somente JSON válido: {"scores":{"q6":0,"q7":0,"q8":0,"q9":0,"q10":0},"essayScore":0,"feedback":"devolutiva curta e específica"}. Não entregue a resposta correta completa; indique o próximo ponto de aprendizagem.`;
  const url='https://generativelanguage.googleapis.com/v1beta/models/'+GEMINI_MODEL+':generateContent?key='+encodeURIComponent(key);
  const body={contents:[{role:'user',parts:[{text:prompt}]}],generationConfig:{temperature:0.1,responseMimeType:'application/json'}};
  const r=UrlFetchApp.fetch(url,{method:'post',contentType:'application/json',payload:JSON.stringify(body),muteHttpExceptions:true});
  if(r.getResponseCode()<200||r.getResponseCode()>=300) throw new Error('Erro na IA ('+r.getResponseCode()+'): '+r.getContentText().slice(0,300));
  const raw=JSON.parse(r.getContentText()), text=(raw.candidates&&raw.candidates[0]&&raw.candidates[0].content&&raw.candidates[0].content.parts||[]).map(p=>p.text||'').join('').replace(/^```json\s*/i,'').replace(/```$/i,'').trim();
  const out=JSON.parse(text), scores=out.scores||{};
  ['q6','q7','q8','q9','q10'].forEach(k=>scores[k]=Math.max(0,Math.min(2,Number(scores[k])||0)));
  out.scores=scores; out.essayScore=Object.values(scores).reduce((a,b)=>a+b,0); return out;
}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
