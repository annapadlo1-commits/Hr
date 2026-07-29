const FIN={
  VERSION:'2.3.1-REAL-WORLD',START:'START',COSTS:'KOSZTY_PRACOWNIKÓW',BUDGETS:'BUDŻETY_LOKALIZACJI',
  EXTRAS:'DODATKI_I_REGUŁY',FORECAST:'PROGNOZA',AUDIT:'AUDYT_FINANSOWY'
};

function onOpen(){
  SpreadsheetApp.getUi().createMenu('GRAFIK PRO — HR & FINANSE')
    .addItem('Instaluj strukturę','finInstall')
    .addItem('Załaduj dane finansowe DEMO','finLoadDemo')
    .addSeparator().addItem('Przelicz prognozę','finRecalculate')
    .addItem('Sprawdź dane','finValidate').addToUi();
}

function finInstall(){
  const defs={
    [FIN.COSTS]:['PRACOWNIK_ID*','MIESIĄC*','TYP_UMOWY','ETAT','STAWKA_BRUTTO_H','KOSZT_PRACODAWCY_H*','KOSZT_STAŁY_MIESIĘCZNY','LIMIT_NADGODZIN_H','AKTYWNY','UWAGI'],
    [FIN.BUDGETS]:['MIESIĄC*','LOKALIZACJA_ID*','BUDŻET_PŁACOWY*','LIMIT_GODZIN','PRÓG_OSTRZEŻENIA_PROC','REZERWA_PROC','AKTYWNY','WŁAŚCICIEL_BUDŻETU'],
    [FIN.EXTRAS]:['KOD','NAZWA','MNOŻNIK','KWOTA_H','WARUNEK','AKTYWNY'],
    [FIN.FORECAST]:['MIESIĄC','LOKALIZACJA_ID','BUDŻET','KOSZT_PLANU','WYKORZYSTANIE_PROC','PROGNOZA_KOŃCA_MIESIĄCA','ODCHYLENIE','STATUS'],
    [FIN.AUDIT]:['CZAS','UŻYTKOWNIK','AKCJA','SZCZEGÓŁY']
  };
  const ss=SpreadsheetApp.getActive();
  Object.keys(defs).forEach(name=>{let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);const h=defs[name];if(sh.getMaxColumns()<h.length)sh.insertColumnsAfter(sh.getMaxColumns(),h.length-sh.getMaxColumns());sh.getRange(1,1,1,h.length).setValues([h]).setBackground('#3f1d66').setFontColor('#fff').setFontWeight('bold');sh.setFrozenRows(1);sh.setHiddenGridlines(true);sh.autoResizeColumns(1,h.length);});
  finStart_();PropertiesService.getDocumentProperties().setProperties({FIN_VERSION:FIN.VERSION,FIN_UPDATED:new Date().toISOString()});
}

function finStart_(){
  const ss=SpreadsheetApp.getActive();let sh=ss.getSheetByName(FIN.START);if(!sh)sh=ss.insertSheet(FIN.START,0);sh.clear();
  sh.getRange('A1:H2').merge().setValue('GRAFIK PRO — HR I FINANSE').setBackground('#2e1065').setFontColor('#fff').setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A4:H4').merge().setValue('CHRONIONA CENTRALA KOSZTÓW, UMÓW I BUDŻETÓW').setBackground('#ede9fe').setFontColor('#5b21b6').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A6:H12').merge().setValue('Ten plik powinien być dostępny wyłącznie dla HR, księgowości i administratora.\n\nPlanner pobiera wyliczone koszty i limity budżetowe. Kierownik nie musi otrzymywać bezpośredniego dostępu do indywidualnych stawek.\n\nPo zmianie danych wybierz „Sprawdź dane”, a następnie zsynchronizuj centrale w Plannerze.').setWrap(true).setVerticalAlignment('middle').setBackground('#faf5ff').setFontSize(13);
  sh.setColumnWidths(1,8,145);sh.setHiddenGridlines(true);
}

function finLoadDemo(){
  finInstall();const month=Utilities.formatDate(new Date(),'Europe/Warsaw','yyyy-MM'),costs=[];
  for(let i=0;i<76;i++){const full=!([27,26,25,46,45,44,64,63,62,71,70,69,75,74,73].includes(i));costs.push([`P${String(i+1).padStart(3,'0')}`,month,full?'UMOWA O PRACĘ':'CZĘŚĆ ETATU',full?1:0.5,30+(i%8),42+(i%9),full?450:180,full?16:8,'TAK',i===0||i===28||i===29||i===47||i===65||i===72?'Menadżer zespołu':'']);}
  finWrite_(FIN.COSTS,costs);
  finWrite_(FIN.BUDGETS,[[month,'KRUCZA',190000,5200,85,8,'TAK','ksiegowosc@demo.pl'],[month,'PAWILONY',45000,1200,85,8,'TAK','ksiegowosc@demo.pl']]);
  finWrite_(FIN.EXTRAS,[['WEEKEND','Dodatek weekendowy',1.25,0,'Sobota lub niedziela','TAK'],['POPOŁUDNIE','Dodatek popołudniowy',1.08,0,'Zmiana POPOŁUDNIE','TAK'],['NOC','Dodatek nocny',1.20,0,'Godziny 22:00–06:00','TAK'],['ŚWIĘTO','Dodatek świąteczny',2.00,0,'Dzień ustawowo wolny','TAK'],['STANDBY','Dyżur stand-by',1,0,'Płatne 2 godziny','TAK'],['NADGODZINY','Nadgodziny',1.50,0,'Ponad nominał','TAK']]);
  finRecalculate();finAudit_('LOAD_DEMO','76 pracowników, 2 budżety');return {ok:true,month};
}

function finWrite_(name,rows){const sh=SpreadsheetApp.getActive().getSheetByName(name);if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();if(rows.length)sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);}

function finValidate(){
  const costs=finRows_(FIN.COSTS),budgets=finRows_(FIN.BUDGETS),errors=[],ids=new Set();
  costs.forEach((r,i)=>{if(!r['PRACOWNIK_ID*']||!r['MIESIĄC*']||r['KOSZT_PRACODAWCY_H*']==='')errors.push(`Koszty wiersz ${i+2}: brak pola obowiązkowego`);if(ids.has(r['PRACOWNIK_ID*']))errors.push(`Duplikat ${r['PRACOWNIK_ID*']}`);ids.add(r['PRACOWNIK_ID*']);});
  budgets.forEach((r,i)=>{if(Number(r['BUDŻET_PŁACOWY*'])<=0)errors.push(`Budżet wiersz ${i+2}: kwota musi być dodatnia`);});
  PropertiesService.getDocumentProperties().setProperty('FIN_UPDATED',new Date().toISOString());finAudit_('VALIDATE',`${errors.length} błędów`);
  SpreadsheetApp.getActive().toast(errors.length?`Błędy: ${errors.length}`:'Dane finansowe poprawne.','Walidacja',6);return {ok:!errors.length,errors};
}

function finRecalculate(){
  const budgets=finRows_(FIN.BUDGETS),month=Utilities.formatDate(new Date(),'Europe/Warsaw','yyyy-MM');
  const rows=budgets.filter(b=>String(b['MIESIĄC*']).slice(0,7)===month).map(b=>[month,b['LOKALIZACJA_ID*'],b['BUDŻET_PŁACOWY*'],0,0,0,Number(b['BUDŻET_PŁACOWY*']),'OCZEKUJE NA PLAN']);
  finWrite_(FIN.FORECAST,rows);return rows;
}

function finRows_(name){const v=SpreadsheetApp.getActive().getSheetByName(name).getDataRange().getValues(),h=v.shift();return v.filter(r=>r.some(Boolean)).map(r=>Object.fromEntries(h.map((x,i)=>[x,r[i]])));}
function finAudit_(action,details){SpreadsheetApp.getActive().getSheetByName(FIN.AUDIT).appendRow([new Date(),Session.getActiveUser().getEmail(),action,details]);}
