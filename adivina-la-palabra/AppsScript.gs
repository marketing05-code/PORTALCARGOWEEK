// CARGOWEEK · Adivina la Palabra
// Reemplaza el código de Apps Script de ESTE juego por este archivo.
// Después: Implementar > Administrar implementaciones > Editar > Nueva versión > Implementar.

const CW_HEADERS = ["Fecha", "Nombre", "Sucursal", "Fase", "Aciertos", "Errores", "Tiempo", "TiempoMs", "Respuesta 1", "Resultado 1", "Correcta 1", "Respuesta 2", "Resultado 2", "Correcta 2", "Respuesta 3", "Resultado 3", "Correcta 3", "Respuesta 4", "Resultado 4", "Correcta 4", "Respuesta 5", "Resultado 5", "Correcta 5", "Respuesta 6", "Resultado 6", "Correcta 6", "Respuesta 7", "Resultado 7", "Correcta 7", "Respuesta 8", "Resultado 8", "Correcta 8", "Respuesta 9", "Resultado 9", "Correcta 9", "Respuesta 10", "Resultado 10", "Correcta 10"];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || "{}");
    if (data.tipo === "progreso") {
      saveLive_(data);
      return json_({ok:true, live:true});
    }

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    ensureHeaders_(sh);

    const record = {
      "Fecha": new Date(),
      "Nombre": data.nombre || "",
      "Sucursal": data.sucursal || "",
      "Fase": data.fase || "",
      "Aciertos": data.aciertos ?? "",
      "Errores": data.errores ?? "",
      "Tiempo": data.tiempo || "",
      "TiempoMs": data.tiempoMs ?? "",
      "Respuesta 1": data["respuesta1"] || (data.respuestas && data.respuestas[0]) || "",
      "Resultado 1": data["resultado1"] || (data.resultados && data.resultados[0]) || "",
      "Correcta 1": data["correcta1"] || (data.detalleRespuestas && data.detalleRespuestas[0] && data.detalleRespuestas[0].respuestaCorrecta) || "",
      "Respuesta 2": data["respuesta2"] || (data.respuestas && data.respuestas[1]) || "",
      "Resultado 2": data["resultado2"] || (data.resultados && data.resultados[1]) || "",
      "Correcta 2": data["correcta2"] || (data.detalleRespuestas && data.detalleRespuestas[1] && data.detalleRespuestas[1].respuestaCorrecta) || "",
      "Respuesta 3": data["respuesta3"] || (data.respuestas && data.respuestas[2]) || "",
      "Resultado 3": data["resultado3"] || (data.resultados && data.resultados[2]) || "",
      "Correcta 3": data["correcta3"] || (data.detalleRespuestas && data.detalleRespuestas[2] && data.detalleRespuestas[2].respuestaCorrecta) || "",
      "Respuesta 4": data["respuesta4"] || (data.respuestas && data.respuestas[3]) || "",
      "Resultado 4": data["resultado4"] || (data.resultados && data.resultados[3]) || "",
      "Correcta 4": data["correcta4"] || (data.detalleRespuestas && data.detalleRespuestas[3] && data.detalleRespuestas[3].respuestaCorrecta) || "",
      "Respuesta 5": data["respuesta5"] || (data.respuestas && data.respuestas[4]) || "",
      "Resultado 5": data["resultado5"] || (data.resultados && data.resultados[4]) || "",
      "Correcta 5": data["correcta5"] || (data.detalleRespuestas && data.detalleRespuestas[4] && data.detalleRespuestas[4].respuestaCorrecta) || "",
      "Respuesta 6": data["respuesta6"] || (data.respuestas && data.respuestas[5]) || "",
      "Resultado 6": data["resultado6"] || (data.resultados && data.resultados[5]) || "",
      "Correcta 6": data["correcta6"] || (data.detalleRespuestas && data.detalleRespuestas[5] && data.detalleRespuestas[5].respuestaCorrecta) || "",
      "Respuesta 7": data["respuesta7"] || (data.respuestas && data.respuestas[6]) || "",
      "Resultado 7": data["resultado7"] || (data.resultados && data.resultados[6]) || "",
      "Correcta 7": data["correcta7"] || (data.detalleRespuestas && data.detalleRespuestas[6] && data.detalleRespuestas[6].respuestaCorrecta) || "",
      "Respuesta 8": data["respuesta8"] || (data.respuestas && data.respuestas[7]) || "",
      "Resultado 8": data["resultado8"] || (data.resultados && data.resultados[7]) || "",
      "Correcta 8": data["correcta8"] || (data.detalleRespuestas && data.detalleRespuestas[7] && data.detalleRespuestas[7].respuestaCorrecta) || "",
      "Respuesta 9": data["respuesta9"] || (data.respuestas && data.respuestas[8]) || "",
      "Resultado 9": data["resultado9"] || (data.resultados && data.resultados[8]) || "",
      "Correcta 9": data["correcta9"] || (data.detalleRespuestas && data.detalleRespuestas[8] && data.detalleRespuestas[8].respuestaCorrecta) || "",
      "Respuesta 10": data["respuesta10"] || (data.respuestas && data.respuestas[9]) || "",
      "Resultado 10": data["resultado10"] || (data.resultados && data.resultados[9]) || "",
      "Correcta 10": data["correcta10"] || (data.detalleRespuestas && data.detalleRespuestas[9] && data.detalleRespuestas[9].respuestaCorrecta) || ""
    };
    appendByHeaders_(sh, record);

    if (data.nombre && data.sucursal) deleteLive_(data.nombre, data.sucursal, data.fase || "");
    return json_({ok:true});
  } catch(err) {
    return json_({ok:false,error:String(err)});
  }
}

function doGet(e) {
  try {
    const action=(e.parameter.action||"").toLowerCase();
    if(action!=="ranking") return output_(e, {ok:true});
    const fase=e.parameter.fase||"";
    const ranking=getRanking_(fase);
    return output_(e, {ok:true,ranking:ranking,updated:new Date().toISOString()});
  } catch(err) {
    return output_(e, {ok:false,error:String(err),ranking:[]});
  }
}

function getRanking_(fase) {
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  ensureHeaders_(sh);
  const vals=sh.getDataRange().getValues();
  const h=vals.length?vals[0].map(String):[];
  const idx={}; h.forEach((x,i)=>idx[normHeader_(x)]=i);
  const finals=[];
  for(let r=1;r<vals.length;r++){
    const row=vals[r];
    const rf=cell_(row,idx,"fase");
    if(fase && rf && String(rf)!==String(fase)) continue;
    const nombre=cell_(row,idx,"nombre"); if(!nombre) continue;
    const sucursal=cell_(row,idx,"sucursal");
    const valor=Number(cell_(row,idx,"aciertos"))||0; const detalle=valor+" aciertos";
    finals.push({nombre:String(nombre),sucursal:String(sucursal||""),valor:Number(valor)||0,detalle:String(detalle||""),estado:"FINAL"});
  }

  // Mantener solo el mejor resultado final de cada persona/sucursal.
  const best={};
  finals.forEach(x=>{
    const k=norm_(x.nombre)+"::"+norm_(x.sucursal);
    if(!best[k] || better_(x,best[k])) best[k]=x;
  });

  // Agregar participantes en vivo actualizados en los últimos 2 minutos.
  const props=PropertiesService.getScriptProperties().getProperties();
  const now=Date.now();
  Object.keys(props).filter(k=>k.indexOf("CW_LIVE_")===0).forEach(k=>{
    try {
      const x=JSON.parse(props[k]);
      if(now-Number(x.actualizado||0)>120000){ PropertiesService.getScriptProperties().deleteProperty(k); return; }
      if(fase && x.fase!==fase) return;
      const pk=norm_(x.nombre)+"::"+norm_(x.sucursal);
      if(best[pk]) return; // si ya terminó, mostrar resultado final
      best[pk]={nombre:x.nombre,sucursal:x.sucursal,valor:Number(x.valor)||0,detalle:x.detalle||"",estado:"EN JUEGO"};
    } catch(_e) {}
  });

  const arr=Object.keys(best).map(k=>best[k]);
  arr.sort((a,b)=>compare_(a,b));
  return arr.slice(0,20);
}

function better_(a,b) {
  return compare_(a,b)<0;
}
function compare_(a,b) {
  // Los resultados finales tienen prioridad sobre los estados en vivo cuando aplica.
  
  return (b.valor-a.valor) || (a.estado==="FINAL"?-1:1);
}

function saveLive_(d) {
  if(!d.nombre||!d.sucursal) return;
  const k=liveKey_(d.nombre,d.sucursal,d.fase||"");
  PropertiesService.getScriptProperties().setProperty(k,JSON.stringify({
    nombre:String(d.nombre),sucursal:String(d.sucursal),fase:String(d.fase||""),
    valor:Number(d.valor)||0,detalle:String(d.detalle||""),actualizado:Date.now()
  }));
}
function deleteLive_(nombre,sucursal,fase) {
  PropertiesService.getScriptProperties().deleteProperty(liveKey_(nombre,sucursal,fase));
}
function liveKey_(n,s,f) {return "CW_LIVE_"+Utilities.base64EncodeWebSafe(norm_(f)+"::"+norm_(n)+"::"+norm_(s));}

function ensureHeaders_(sh) {
  const last=Math.max(sh.getLastColumn(),1);
  let h=sh.getRange(1,1,1,last).getValues()[0].map(String);
  if(sh.getLastRow()===0 || h.every(x=>!x.trim())) {
    sh.getRange(1,1,1,CW_HEADERS.length).setValues([CW_HEADERS]); return;
  }
  const existing=h.map(normHeader_);
  CW_HEADERS.forEach(name=>{
    if(existing.indexOf(normHeader_(name))<0){h.push(name);existing.push(normHeader_(name));}
  });
  sh.getRange(1,1,1,h.length).setValues([h]);
}
function appendByHeaders_(sh, record) {
  const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const row=h.map(name=>{
    const key=Object.keys(record).find(k=>normHeader_(k)===normHeader_(name));
    return key ? record[key] : "";
  });
  sh.appendRow(row);
}
function cell_(row,idx,name) {
  const i=idx[normHeader_(name)]; return i===undefined?"":row[i];
}
function normHeader_(s) {return norm_(s).replace(/ /g,"");}
function norm_(s) {return String(s||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-z0-9]+/g," ").trim();}
function parseTimeMs_(s) {
  const m=String(s||"").match(/(\d+):(\d+)(?:\.(\d+))?/); if(!m)return 999999999;
  return Number(m[1])*60000+Number(m[2])*1000+Number((m[3]||"0").padEnd(2,"0").slice(0,2))*10;
}
function json_(obj) {return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
function output_(e,obj) {
  const cb=e.parameter.callback;
  if(cb) return ContentService.createTextOutput(cb+"("+JSON.stringify(obj)+");").setMimeType(ContentService.MimeType.JAVASCRIPT);
  return json_(obj);
}
