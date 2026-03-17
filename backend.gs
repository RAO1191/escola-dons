// ═══════════════════════════════════════════════════════════════
// ESCOLA DE DONS — Google Apps Script Backend v3
// ═══════════════════════════════════════════════════════════════

const SHEET_ID = '1sfxfi_PBwWCWaW9xUwl9npIfLeUcsjVZwjINPu7R1CU';
const SS = SpreadsheetApp.openById(SHEET_ID);

var AM = {1:"A",23:"A",45:"A",67:"A",89:"A",2:"B",24:"B",46:"B",68:"B",90:"B",3:"C",25:"C",47:"C",69:"C",91:"C",4:"D",26:"D",48:"D",70:"D",92:"D",5:"E",27:"E",49:"E",71:"E",93:"E",6:"F",28:"F",50:"F",72:"F",94:"F",7:"G",29:"G",51:"G",73:"G",95:"G",8:"H",30:"H",52:"H",74:"H",96:"H",9:"I",31:"I",53:"I",75:"I",97:"I",10:"J",32:"J",54:"J",76:"J",98:"J",11:"K",33:"K",55:"K",77:"K",99:"K",12:"L",34:"L",56:"L",78:"L",100:"L",13:"M",35:"M",57:"M",79:"M",101:"M",14:"N",36:"N",58:"N",80:"N",102:"N",15:"O",37:"O",59:"O",81:"O",103:"O",16:"P",38:"P",60:"P",82:"P",104:"P",17:"Q",39:"Q",61:"Q",83:"Q",105:"Q",18:"R",40:"R",62:"R",84:"R",106:"R",19:"S",41:"S",63:"S",85:"S",107:"S",20:"T",42:"T",64:"T",86:"T",108:"T",21:"U",43:"U",65:"U",87:"U",109:"U",22:"V",44:"V",66:"V",88:"V",110:"V"};
var OM = {1:"A",2:"B",3:"C",4:"D",5:"E",6:"F",7:"G",8:"H",9:"I",10:"J",11:"K",12:"L",13:"M",14:"N",15:"O",16:"P",17:"Q",18:"R",19:"S",20:"T",21:"U",22:"V",23:"W",24:"A",25:"B",26:"C",27:"D",28:"E",29:"F",30:"G",31:"H",32:"I",33:"J",34:"K",35:"L",36:"M",37:"N",38:"O",39:"P",40:"Q",41:"R",42:"S",43:"T",44:"U",45:"V",46:"W",47:"A",48:"B",49:"C",50:"D",51:"E",52:"F",53:"G",54:"H",55:"I",56:"J",57:"K",58:"L",59:"M",60:"N",61:"O",62:"P",63:"Q",64:"R",65:"S",66:"T",67:"U",68:"V",69:"W",70:"A",71:"B",72:"C",73:"D",74:"E",75:"F",76:"G",77:"H",78:"I",79:"J",80:"K",81:"L",82:"M",83:"N",84:"O",85:"P",86:"Q",87:"R",88:"S",89:"T",90:"U",91:"V",92:"W",93:"A",94:"B",95:"C",96:"D",97:"E",98:"F",99:"G",100:"H",101:"I",102:"J",103:"K",104:"L",105:"M",106:"N",107:"O",108:"P",109:"Q",110:"R",111:"S",112:"T",113:"U",114:"V",115:"W",116:"A",117:"B",118:"C",119:"D",120:"E",121:"F",122:"G",123:"H",124:"I",125:"J",126:"K",127:"L",128:"M",129:"N",130:"O",131:"P",132:"Q",133:"R",134:"S",135:"T",136:"U",137:"V",138:"W"};
var DON_ALFA = {"A":"Administração","B":"Apostolado","C":"Conhecimento","D":"Contribuição","E":"Crianças","F":"Discernimento","G":"Evangelismo","H":"Exortação","I":"Fé","J":"Hospitalidade","K":"Intercessão","L":"Liderança","M":"Louvor","N":"Ensino","O":"Misericórdia","P":"Missão","Q":"Pastoreio","R":"Profecia","S":"Profeta","T":"Sabedoria","U":"Serviço","V":"Socorros"};
var DON_OMEGA = {"A":"Administração","B":"Apostolado","C":"Discernimento","D":"Evangelismo","E":"Exortação","F":"Fé","G":"Contribuição","H":"Hospitalidade","I":"Conhecimento","J":"Liderança","K":"Misericórdia","L":"Missão","M":"Profecia","N":"Pastoreio","O":"Serviço","P":"Ensino","Q":"Sabedoria","R":"Milagres","S":"Cura","T":"Línguas","U":"Interpretação","V":"Intercessão","W":"Libertação"};
var ALFA_NORM = {"A":"administracao","B":"apostolado","C":"conhecimento","D":"contribuicao","E":"criancas","F":"discernimento","G":"evangelismo","H":"exortacao","I":"fe","J":"hospitalidade","K":"intercessao","L":"lideranca","M":"louvor","N":"ensino","O":"misericordia","P":"missao","Q":"pastoreio","R":"profecia","S":"profeta","T":"sabedoria","U":"servico","V":"socorros"};
var OMEGA_NORM = {"A":"administracao","B":"apostolado","C":"discernimento","D":"evangelismo","E":"exortacao","F":"fe","G":"contribuicao","H":"hospitalidade","I":"conhecimento","J":"lideranca","K":"misericordia","L":"missao","M":"profecia","N":"pastoreio","O":"servico","P":"ensino","Q":"sabedoria","R":"milagres","S":"cura","T":"linguas","U":"interpretacao","V":"intercessao","W":"libertacao"};

function makeResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
function doGet(e) {
  try {
    switch(e.parameter.action) {
      case 'getUser':          return makeResponse(getUser(e.parameter.code));
      case 'getAnswers':       return makeResponse(getAnswers(e.parameter.code));
      case 'getAllUsers':      return makeResponse(getAllUsers(e.parameter.senha));
      case 'getStats':         return makeResponse(getStats(e.parameter.senha));
      case 'getConfig':        return makeResponse(getConfig());
      case 'getBloqueioGeral': return makeResponse(getBloqueioGeral());
      case 'getExportData':    return makeResponse(getExportData(e.parameter.senha));
      case 'getRankingData':   return makeResponse(getRankingData(e.parameter.senha));
      default:                 return makeResponse({ok:false,error:'Ação inválida'});
    }
  } catch(err) { return makeResponse({ok:false,error:err.message}); }
}
function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  try {
    switch(body.action) {
      case 'register':          return makeResponse(register(body));
      case 'saveAnswers':       return makeResponse(saveAnswers(body));
      case 'saveDons':          return makeResponse(saveDons(body));        // ← NOVO
      case 'releaseResult':     return makeResponse(releaseResult(body));
      case 'lockResult':        return makeResponse(lockResult(body));
      case 'releaseMany':       return makeResponse(releaseMany(body));
      case 'blockForm':         return makeResponse(blockForm(body));
      case 'unblockForm':       return makeResponse(unblockForm(body));
      case 'setBloqueioGeral':  return makeResponse(setBloqueioGeral(body));
      case 'deleteUser':        return makeResponse(deleteUser(body));
      case 'uploadImage':       return makeResponse(uploadImage(body));
      case 'saveConfig':        return makeResponse(saveConfig(body));
      default:                  return makeResponse({ok:false,error:'Ação inválida'});
    }
  } catch(err) { return makeResponse({ok:false,error:err.message}); }
}

function checkAdmin(senha) {
  var data = SS.getSheetByName('config').getDataRange().getValues();
  for(var i=1;i<data.length;i++) {
    if(data[i][0]==='senhaAdmin') {
      if(data[i][1].toString()!==senha.toString()) throw new Error('Senha incorreta.');
      return true;
    }
  }
  throw new Error('Senha não configurada.');
}
function getUserRow(sheet, code) {
  var data = sheet.getDataRange().getValues();
  for(var i=1;i<data.length;i++) if(data[i][0]===code) return {rowIndex:i+1,row:data[i]};
  return {rowIndex:0,row:null};
}
function generateCode(sheet) {
  var L='ABCDEFGHJKLMNPQRSTUVWXYZ', D='0123456789';
  var existing = new Set(sheet.getDataRange().getValues().slice(1).map(function(r){return r[0];}));
  var code;
  do {
    code='';
    for(var i=0;i<3;i++) code+=L[Math.floor(Math.random()*L.length)];
    for(var i=0;i<3;i++) code+=D[Math.floor(Math.random()*D.length)];
  } while(existing.has(code));
  return code;
}

function calcTotalsGS(ans, map) {
  var t = {};
  Object.keys(ans).forEach(function(q) {
    var l = map[parseInt(q)];
    if(l) t[l] = (t[l]||0) + parseInt(ans[q]);
  });
  return t;
}
function getTop3GS(totals) {
  var sorted = Object.keys(totals).map(function(l){return [l,totals[l]];}).sort(function(a,b){return b[1]-a[1];});
  var groups = [];
  sorted.forEach(function(x) {
    var last = groups[groups.length-1];
    if(last && last.score===x[1]) last.letters.push(x[0]);
    else groups.push({score:x[1], letters:[x[0]]});
  });
  return groups.slice(0,3);
}
function calcCombinationGS(alfaGroups, omegaGroups) {
  var alfaNames = {};
  alfaGroups.forEach(function(g) { g.letters.forEach(function(l) { alfaNames[ALFA_NORM[l]||l] = DON_ALFA[l]||l; }); });
  var omegaNames = {};
  omegaGroups.forEach(function(g) { g.letters.forEach(function(l) { omegaNames[OMEGA_NORM[l]||l] = DON_OMEGA[l]||l; }); });
  var primary = [], secondary = [];
  Object.keys(alfaNames).forEach(function(n) {
    if(omegaNames[n]) primary.push(alfaNames[n]);
    else secondary.push(alfaNames[n]);
  });
  Object.keys(omegaNames).forEach(function(n) { if(!alfaNames[n]) secondary.push(omegaNames[n]); });
  return {primary:primary, secondary:secondary};
}

function getOrCreateSheet(name, headers) {
  var sh = SS.getSheetByName(name);
  if(!sh) {
    sh = SS.insertSheet(name);
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function upsertRow(sheet, code, newRow) {
  var data = sheet.getDataRange().getValues();
  for(var i=1;i<data.length;i++) { if(data[i][0]===code) { sheet.getRange(i+1,1,1,newRow.length).setValues([newRow]); return; } }
  sheet.appendRow(newRow);
}
function saveResultSheets(code, nome, alfa, omega, now) {
  var tA=calcTotalsGS(alfa,AM), tO=calcTotalsGS(omega,OM);
  var gA=getTop3GS(tA), gO=getTop3GS(tO);
  var comb=calcCombinationGS(gA,gO);
  var shA = getOrCreateSheet('Resultado Alfa', ['Código','Nome','Data','1º Lugar - Letra(s)','1º Lugar - Dom(ns)','1º Pts','2º Lugar - Letra(s)','2º Lugar - Dom(ns)','2º Pts','3º Lugar - Letra(s)','3º Lugar - Dom(ns)','3º Pts']);
  var rowA = [code,nome,now];
  for(var i=0;i<3;i++) { if(gA[i]){rowA.push(gA[i].letters.join(' · '),gA[i].letters.map(function(l){return DON_ALFA[l]||l;}).join(', '),gA[i].score);}else{rowA.push('','',0);} }
  upsertRow(shA,code,rowA);
  var shO = getOrCreateSheet('Resultado Ômega', ['Código','Nome','Data','1º Lugar - Letra(s)','1º Lugar - Dom(ns)','1º Pts','2º Lugar - Letra(s)','2º Lugar - Dom(ns)','2º Pts','3º Lugar - Letra(s)','3º Lugar - Dom(ns)','3º Pts']);
  var rowO = [code,nome,now];
  for(var i=0;i<3;i++) { if(gO[i]){rowO.push(gO[i].letters.join(' · '),gO[i].letters.map(function(l){return DON_OMEGA[l]||l;}).join(', '),gO[i].score);}else{rowO.push('','',0);} }
  upsertRow(shO,code,rowO);
  var shC = getOrCreateSheet('Comparativo', ['Código','Nome','Data','Dons Primários','Dons Secundários']);
  upsertRow(shC,code,[code,nome,now,comb.primary.join(', '),comb.secondary.join(', ')]);
}

// ── REGISTER ──────────────────────────────────────────────────
function register(body) {
  var nome=body.nome, telefone=body.telefone;
  if(!nome||!telefone) throw new Error('Nome e telefone são obrigatórios.');
  var cfg=getConfig();
  if(cfg.bloqueioGeral==='Sim') return {ok:false,error:'O sistema está bloqueado. Novos cadastros não são permitidos no momento.'};
  var sheet=SS.getSheetByName('usuarios');
  var code=generateCode(sheet);
  var now=new Date().toLocaleString('pt-BR',{timeZone:'America/Sao_Paulo'});
  // col: code,nome,telefone,dataCadastro,alfaResp,omegaResp,completo,resultadoLib,don1,don2,don3,formBloqueado
  sheet.appendRow([code,nome,telefone,now,0,0,'Não','Não','','','','Não']);
  return {ok:true,code:code,nome:nome};
}

// ── HELPERS: duas tabelas de respostas ───────────────────────
function getResSheet(inv) {
  var name = inv==='alfa' ? 'respostas_alfa' : 'respostas_omega';
  var sh = SS.getSheetByName(name);
  if(!sh) {
    sh = SS.insertSheet(name);
    sh.getRange(1,1,1,5).setValues([['codigo','inventario','questao','valor','data']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── SAVE ANSWERS ──────────────────────────────────────────────
function saveAnswers(body) {
  var code=body.code, inv=body.inv;
  if(!code) throw new Error('Código inválido.');
  var usrSheet=SS.getSheetByName('usuarios');
  var ud=getUserRow(usrSheet,code);
  if(ud.rowIndex>0 && ud.row[11]==='Sim') return {ok:false,error:'Preenchimento bloqueado pelo administrador.'};
  var resSheet=getResSheet(inv);
  var now=new Date().toLocaleString('pt-BR',{timeZone:'America/Sao_Paulo'});
  var allData=resSheet.getDataRange().getValues();
  var rowsToClear=[];
  for(var i=allData.length-1;i>=1;i--) {
    if(allData[i][0]===code) rowsToClear.push(i+1);
  }
  if(rowsToClear.length>0) {
    rowsToClear.sort(function(a,b){return b-a;});
    var groups=[], cur=[rowsToClear[0]];
    for(var j=1;j<rowsToClear.length;j++) {
      if(rowsToClear[j]===cur[cur.length-1]-1) cur.push(rowsToClear[j]);
      else { groups.push(cur); cur=[rowsToClear[j]]; }
    }
    groups.push(cur);
    groups.forEach(function(g) {
      resSheet.deleteRows(g[g.length-1], g.length);
    });
  }
  var answers=body.answers||{};
  var newRows=[];
  Object.keys(answers).forEach(function(q){
    newRows.push([code,inv,parseInt(q),parseInt(answers[q]),now]);
  });
  if(newRows.length>0) resSheet.getRange(resSheet.getLastRow()+1,1,newRows.length,5).setValues(newRows);
  var ansCount=newRows.length;
  var colIdx=inv==='alfa'?5:6;
  if(ud.rowIndex>0) usrSheet.getRange(ud.rowIndex,colIdx).setValue(ansCount);
  var alfaCount=inv==='alfa'?ansCount:parseInt(ud.row[4])||0;
  var omegaCount=inv==='omega'?ansCount:parseInt(ud.row[5])||0;
  var completo=alfaCount>=110&&omegaCount>=138?'Sim':'Não';
  if(ud.rowIndex>0) usrSheet.getRange(ud.rowIndex,7).setValue(completo);
  return {ok:true,completo:completo};
}

// ── SAVE DONS (chamado pelo form após calcular resultado) ─────  ← NOVO
function saveDons(body) {
  var sheet = SS.getSheetByName('usuarios');
  var ud = getUserRow(sheet, body.code);
  if(ud.rowIndex === 0) return {ok:false, error:'Usuário não encontrado.'};
  sheet.getRange(ud.rowIndex, 9).setValue(body.don1 || '');
  sheet.getRange(ud.rowIndex, 10).setValue(body.don2 || '');
  sheet.getRange(ud.rowIndex, 11).setValue(body.don3 || '');
  return {ok:true};
}

// ── RELEASE RESULT (grava nas 3 abas ao liberar) ──────────────
function releaseAndSaveSheets(body) {
  checkAdmin(body.senha);
  var usrSheet=SS.getSheetByName('usuarios');
  var ud=getUserRow(usrSheet,body.code);
  if(ud.rowIndex===0) throw new Error('Usuário não encontrado.');
  var alfa={}, omega={};
  var shA=SS.getSheetByName('respostas_alfa'), shO=SS.getSheetByName('respostas_omega');
  if(shA){var dA=shA.getDataRange().getValues();for(var i=1;i<dA.length;i++){if(dA[i][0]===body.code) alfa[dA[i][2]]=dA[i][3];}}
  if(shO){var dO=shO.getDataRange().getValues();for(var i=1;i<dO.length;i++){if(dO[i][0]===body.code) omega[dO[i][2]]=dO[i][3];}}
  var now=new Date().toLocaleString('pt-BR',{timeZone:'America/Sao_Paulo'});
  saveResultSheets(body.code, ud.row[1], alfa, omega, now);

  // ← NOVO: salva top 3 dons Alfa na aba usuarios
  var tA = calcTotalsGS(alfa, AM);
  var gA = getTop3GS(tA);
  usrSheet.getRange(ud.rowIndex, 9).setValue(gA[0] ? gA[0].letters[0] : '');
  usrSheet.getRange(ud.rowIndex, 10).setValue(gA[1] ? gA[1].letters[0] : '');
  usrSheet.getRange(ud.rowIndex, 11).setValue(gA[2] ? gA[2].letters[0] : '');

  usrSheet.getRange(ud.rowIndex,8).setValue('Sim');
  return {ok:true};
}

// ── BLOCK / UNBLOCK FORM ──────────────────────────────────────
function blockForm(body) {
  checkAdmin(body.senha);
  var sheet=SS.getSheetByName('usuarios'), ud=getUserRow(sheet,body.code);
  if(ud.rowIndex===0) throw new Error('Usuário não encontrado.');
  sheet.getRange(ud.rowIndex,12).setValue('Sim');
  return {ok:true};
}
function unblockForm(body) {
  checkAdmin(body.senha);
  var sheet=SS.getSheetByName('usuarios'), ud=getUserRow(sheet,body.code);
  if(ud.rowIndex===0) throw new Error('Usuário não encontrado.');
  sheet.getRange(ud.rowIndex,12).setValue('Não');
  return {ok:true};
}

// ── GET USER ──────────────────────────────────────────────────
function getUser(code) {
  if(!code) return {ok:false,error:'Código inválido.'};
  var usrSheet=SS.getSheetByName('usuarios');
  var ud=getUserRow(usrSheet,code);
  if(ud.rowIndex===0) return {ok:false,error:'Usuário não encontrado.'};
  var row=ud.row;
  var alfa={}, omega={};
  var shA=SS.getSheetByName('respostas_alfa'), shO=SS.getSheetByName('respostas_omega');
  if(shA){var dA=shA.getDataRange().getValues();for(var i=1;i<dA.length;i++){if(dA[i][0]===code) alfa[dA[i][2]]=dA[i][3];}}
  if(shO){var dO=shO.getDataRange().getValues();for(var i=1;i<dO.length;i++){if(dO[i][0]===code) omega[dO[i][2]]=dO[i][3];}}
  return {ok:true,code:row[0],nome:row[1],telefone:row[2],dataCadastro:row[3],alfaRespondidas:row[4],omegaRespondidas:row[5],completo:row[6],resultadoLiberado:row[7]==='Sim',don1:row[8],don2:row[9],don3:row[10],preenchimentoBloqueado:row[11]==='Sim',alfa:alfa,omega:omega};
}

// ── GET ANSWERS ───────────────────────────────────────────────
function getAnswers(code) {
  if(!code) return {ok:false,error:'Código inválido.'};
  var alfa={}, omega={};
  var shA=SS.getSheetByName('respostas_alfa'), shO=SS.getSheetByName('respostas_omega');
  if(shA){var dA=shA.getDataRange().getValues();for(var i=1;i<dA.length;i++){if(dA[i][0]===code) alfa[dA[i][2]]=dA[i][3];}}
  if(shO){var dO=shO.getDataRange().getValues();for(var i=1;i<dO.length;i++){if(dO[i][0]===code) omega[dO[i][2]]=dO[i][3];}}
  return {ok:true,alfa:alfa,omega:omega};
}

// ── GET ALL USERS ─────────────────────────────────────────────
function getAllUsers(senha) {
  checkAdmin(senha);
  var data=SS.getSheetByName('usuarios').getDataRange().getValues();
  var users=[];
  for(var i=1;i<data.length;i++) {
    if(!data[i][0]) continue;
    var row=data[i];
    users.push({code:row[0],nome:row[1],telefone:row[2],dataCadastro:row[3],alfaRespondidas:row[4],omegaRespondidas:row[5],completo:row[6],resultadoLiberado:row[7]==='Sim',don1:row[8],don2:row[9],don3:row[10],preenchimentoBloqueado:row[11]==='Sim'});
  }
  return {ok:true,users:users};
}

// ── GET STATS ─────────────────────────────────────────────────
function getStats(senha) {
  checkAdmin(senha);
  var data=SS.getSheetByName('usuarios').getDataRange().getValues();
  var donCount={},total=0,completos=0,liberados=0,parciais=0;
  var shA=SS.getSheetByName('respostas_alfa'), shO=SS.getSheetByName('respostas_omega');
  var alfaAll={}, omegaAll={};
  if(shA){var dA=shA.getDataRange().getValues();for(var i=1;i<dA.length;i++){if(!alfaAll[dA[i][0]])alfaAll[dA[i][0]]={};alfaAll[dA[i][0]][dA[i][2]]=dA[i][3];}}
  if(shO){var dO=shO.getDataRange().getValues();for(var i=1;i<dO.length;i++){if(!omegaAll[dO[i][0]])omegaAll[dO[i][0]]={};omegaAll[dO[i][0]][dO[i][2]]=dO[i][3];}}
  for(var i=1;i<data.length;i++) {
    if(!data[i][0]) continue;
    total++;
    var a=parseInt(data[i][4])||0, o=parseInt(data[i][5])||0;
    if(a>=110&&o>=138) completos++; else if(a>0||o>0) parciais++;
    if(data[i][7]==='Sim') liberados++;
    var code=data[i][0];
    var alfa=alfaAll[code]||{}, omega=omegaAll[code]||{};
    if(Object.keys(alfa).length>0||Object.keys(omega).length>0) {
      var tA=calcTotalsGS(alfa,AM), tO=calcTotalsGS(omega,OM);
      var gA=getTop3GS(tA), gO=getTop3GS(tO);
      var comb=calcCombinationGS(gA,gO);
      comb.primary.concat(comb.secondary).slice(0,3).forEach(function(d){if(d)donCount[d]=(donCount[d]||0)+1;});
    }
  }
  return {ok:true,total:total,completos:completos,parciais:parciais,liberados:liberados,donCount:donCount};
}

// ── RELEASE / LOCK RESULTADO ──────────────────────────────────
function releaseResult(body) {
  return releaseAndSaveSheets(body);
}
function lockResult(body) {
  checkAdmin(body.senha);
  var sheet=SS.getSheetByName('usuarios'), ud=getUserRow(sheet,body.code);
  if(ud.rowIndex===0) throw new Error('Usuário não encontrado.');
  sheet.getRange(ud.rowIndex,8).setValue('Não');
  return {ok:true};
}
function releaseMany(body) {
  checkAdmin(body.senha);
  var sheet=SS.getSheetByName('usuarios');
  body.codes.forEach(function(code){ var ud=getUserRow(sheet,code); if(ud.rowIndex>0) sheet.getRange(ud.rowIndex,8).setValue('Sim'); });
  return {ok:true,count:body.codes.length};
}

// ── DELETE USER ───────────────────────────────────────────────
function deleteUser(body) {
  checkAdmin(body.senha);
  var usrSheet=SS.getSheetByName('usuarios');
  var ud=getUserRow(usrSheet,body.code);
  if(ud.rowIndex>0) usrSheet.deleteRow(ud.rowIndex);
  ['respostas_alfa','respostas_omega'].forEach(function(shName){
    var resSheet=SS.getSheetByName(shName); if(!resSheet) return;
    var resData=resSheet.getDataRange().getValues();
    var toDelete=[];
    for(var i=resData.length-1;i>=1;i--) if(resData[i][0]===body.code) toDelete.push(i+1);
    if(toDelete.length>0) {
      toDelete.sort(function(a,b){return b-a;});
      var groups=[], cur=[toDelete[0]];
      for(var j=1;j<toDelete.length;j++) {
        if(toDelete[j]===cur[cur.length-1]-1) cur.push(toDelete[j]);
        else { groups.push(cur); cur=[toDelete[j]]; }
      }
      groups.push(cur);
      groups.forEach(function(g){ resSheet.deleteRows(g[g.length-1], g.length); });
    }
  });
  ['Resultado Alfa','Resultado Omega','Comparativo'].forEach(function(name){
    var sh=SS.getSheetByName(name); if(!sh) return;
    var d=sh.getDataRange().getValues();
    for(var i=d.length-1;i>=1;i--) if(d[i][0]===body.code) sh.deleteRow(i+1);
  });
  return {ok:true};
}

// ── BLOQUEIO GERAL ────────────────────────────────────────────
function getBloqueioGeral() {
  try {
    var cfg = getConfig();
    return {ok:true, bloqueado: cfg.bloqueioGeral==='Sim'};
  } catch(e) {
    return {ok:true, bloqueado: false};
  }
}
function setBloqueioGeral(body) {
  checkAdmin(body.senha);
  var sheet=SS.getSheetByName('config');
  var data=sheet.getDataRange().getValues();
  var found=false;
  for(var i=1;i<data.length;i++){if(data[i][0]==='bloqueioGeral'){sheet.getRange(i+1,2).setValue(body.bloqueado?'Sim':'Não');found=true;break;}}
  if(!found) sheet.appendRow(['bloqueioGeral', body.bloqueado?'Sim':'Não']);
  return {ok:true};
}

// ── EXPORT EXCEL DATA ─────────────────────────────────────────
function getExportData(senha) {
  checkAdmin(senha);
  var users = SS.getSheetByName('usuarios').getDataRange().getValues();
  var shA=SS.getSheetByName('respostas_alfa'), shO=SS.getSheetByName('respostas_omega');
  var alfaAll={}, omegaAll={};
  if(shA){var dA=shA.getDataRange().getValues();for(var i=1;i<dA.length;i++){if(!alfaAll[dA[i][0]])alfaAll[dA[i][0]]={};alfaAll[dA[i][0]][dA[i][2]]=dA[i][3];}}
  if(shO){var dO=shO.getDataRange().getValues();for(var i=1;i<dO.length;i++){if(!omegaAll[dO[i][0]])omegaAll[dO[i][0]]={};omegaAll[dO[i][0]][dO[i][2]]=dO[i][3];}}
  var rows=[];
  for(var i=1;i<users.length;i++){
    var row=users[i]; if(!row[0]) continue;
    var code=row[0], nome=row[1];
    var alfa=alfaAll[code]||{}, omega=omegaAll[code]||{};
    var tA=calcTotalsGS(alfa,AM), tO=calcTotalsGS(omega,OM);
    var gA=getTop3GS(tA), gO=getTop3GS(tO);
    var comb=calcCombinationGS(gA,gO);
    var getTop=function(g,nameMap,n){if(!g[n])return'';return g[n].letters.map(function(l){return nameMap[l]||l;}).join(', ')+' ('+g[n].score+'pts)';};
    rows.push({
      codigo:code, nome:nome,
      alfa1:getTop(gA,DON_ALFA,0), alfa2:getTop(gA,DON_ALFA,1), alfa3:getTop(gA,DON_ALFA,2),
      omega1:getTop(gO,DON_OMEGA,0), omega2:getTop(gO,DON_OMEGA,1), omega3:getTop(gO,DON_OMEGA,2),
      primarios:comb.primary.join(', '), secundarios:comb.secondary.join(', ')
    });
  }
  return {ok:true, rows:rows};
}

// ── RANKING DATA ──────────────────────────────────────────────
function getRankingData(senha) {
  checkAdmin(senha);
  var users = SS.getSheetByName('usuarios').getDataRange().getValues();
  var shA=SS.getSheetByName('respostas_alfa'), shO=SS.getSheetByName('respostas_omega');
  var alfaAll={}, omegaAll={}, nomeMap={};
  for(var i=1;i<users.length;i++){if(users[i][0])nomeMap[users[i][0]]=users[i][1];}
  if(shA){var dA=shA.getDataRange().getValues();for(var i=1;i<dA.length;i++){if(!alfaAll[dA[i][0]])alfaAll[dA[i][0]]={};alfaAll[dA[i][0]][dA[i][2]]=dA[i][3];}}
  if(shO){var dO=shO.getDataRange().getValues();for(var i=1;i<dO.length;i++){if(!omegaAll[dO[i][0]])omegaAll[dO[i][0]]={};omegaAll[dO[i][0]][dO[i][2]]=dO[i][3];}}
  var alfaRank={}, omegaRank={}, combRank={};
  Object.keys(nomeMap).forEach(function(code){
    var alfa=alfaAll[code]||{}, omega=omegaAll[code]||{};
    var tA=calcTotalsGS(alfa,AM), tO=calcTotalsGS(omega,OM);
    var gA=getTop3GS(tA), gO=getTop3GS(tO), comb=calcCombinationGS(gA,gO);
    Object.keys(tA).forEach(function(l){if(!alfaRank[l])alfaRank[l]=[];alfaRank[l].push({code:code,nome:nomeMap[code],pts:tA[l]});});
    Object.keys(tO).forEach(function(l){if(!omegaRank[l])omegaRank[l]=[];omegaRank[l].push({code:code,nome:nomeMap[code],pts:tO[l]});});
    comb.primary.forEach(function(n){if(!combRank[n])combRank[n]=[];
      var ptA=0,ptO=0;
      Object.keys(DON_ALFA).forEach(function(l){if(DON_ALFA[l]===n)ptA=tA[l]||0;});
      Object.keys(DON_OMEGA).forEach(function(l){if(DON_OMEGA[l]===n)ptO=tO[l]||0;});
      combRank[n].push({code:code,nome:nomeMap[code],pts:ptA+ptO});
    });
  });
  // alfaRank e omegaRank: top 10 para o modal de Ranking
  [alfaRank,omegaRank].forEach(function(rank){
    Object.keys(rank).forEach(function(k){rank[k].sort(function(a,b){return b.pts-a.pts;});rank[k]=rank[k].slice(0,10);});
  });
  // combRank: sem limite — usado pelo Relatório por Dons que mostra todos
  Object.keys(combRank).forEach(function(k){combRank[k].sort(function(a,b){return b.pts-a.pts;});});
  return {ok:true, alfa:alfaRank, omega:omegaRank, cruzamento:combRank, donAlfa:DON_ALFA, donOmega:DON_OMEGA};
}

// ── CONFIG ────────────────────────────────────────────────────
function getConfig() {
  var data=SS.getSheetByName('config').getDataRange().getValues();
  var cfg={ok:true};
  for(var i=1;i<data.length;i++) if(data[i][0]) cfg[data[i][0]]=data[i][1];
  return cfg;
}
function saveConfig(body) {
  checkAdmin(body.senha);
  var sheet=SS.getSheetByName('config');
  var data=sheet.getDataRange().getValues();
  var keys=['churchName','bgFileId','logoFileId','senhaAdmin'];
  keys.forEach(function(key) {
    if(body[key]===undefined) return;
    var found=false;
    for(var i=1;i<data.length;i++) { if(data[i][0]===key){sheet.getRange(i+1,2).setValue(body[key]);data[i][1]=body[key];found=true;break;} }
    if(!found){sheet.appendRow([key,body[key]]);data.push([key,body[key]]);}
  });
  return {ok:true};
}

// ── UPLOAD IMAGE ──────────────────────────────────────────────
function uploadImage(body) {
  checkAdmin(body.senha);
  var folderName='Escola de Dons - Imagens', folder;
  var folders=DriveApp.getFoldersByName(folderName);
  if(folders.hasNext()) folder=folders.next(); else folder=DriveApp.createFolder(folderName);
  var oldKey=body.type==='logo'?'logoFileId':'bgFileId';
  var cfg=getConfig();
  if(cfg[oldKey]){try{DriveApp.getFileById(cfg[oldKey]).setTrashed(true);}catch(e){}}
  var blob=Utilities.newBlob(Utilities.base64Decode(body.base64Data),body.mimeType,body.filename);
  var file=folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
  var configSheet=SS.getSheetByName('config');
  var configData=configSheet.getDataRange().getValues();
  var found=false;
  for(var i=1;i<configData.length;i++) { if(configData[i][0]===oldKey){configSheet.getRange(i+1,2).setValue(file.getId());found=true;break;} }
  if(!found) configSheet.appendRow([oldKey,file.getId()]);
  return {ok:true,fileId:file.getId()};
}

// ── COMPUTE TOP3 ──────────────────────────────────────────────
function computeTop3(alfa, omega) {
  var tA=calcTotalsGS(alfa,AM), tO=calcTotalsGS(omega,OM);
  var gA=getTop3GS(tA), gO=getTop3GS(tO);
  var comb=calcCombinationGS(gA,gO);
  return comb.primary.concat(comb.secondary).slice(0,3);
}

// ── BACKFILL DONS (executar UMA VEZ para usuários existentes) ─
function backfillDons() {
  var usrSheet = SS.getSheetByName('usuarios');
  var users = usrSheet.getDataRange().getValues();
  var shA = SS.getSheetByName('respostas_alfa');
  var alfaAll = {};
  if(shA){
    var dA=shA.getDataRange().getValues();
    for(var i=1;i<dA.length;i++){
      if(!alfaAll[dA[i][0]]) alfaAll[dA[i][0]]={};
      alfaAll[dA[i][0]][dA[i][2]]=dA[i][3];
    }
  }
  var count = 0;
  for(var i=1;i<users.length;i++){
    var code=users[i][0]; if(!code) continue;
    if(users[i][8] && users[i][9]) continue; // já tem dons, pula
    var alfa=alfaAll[code]||{};
    if(!Object.keys(alfa).length) continue;
    var tA=calcTotalsGS(alfa,AM);
    var gA=getTop3GS(tA);
    usrSheet.getRange(i+1,9).setValue(gA[0]?gA[0].letters[0]:'');
    usrSheet.getRange(i+1,10).setValue(gA[1]?gA[1].letters[0]:'');
    usrSheet.getRange(i+1,11).setValue(gA[2]?gA[2].letters[0]:'');
    count++;
  }
  Logger.log('Backfill concluído: ' + count + ' usuários atualizados.');
}
