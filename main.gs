/****************************************************
 🎭 CMT MUSICAL - PORTALE COMPLETO (SISTEMA UNIFICATO)
 ****************************************************/

function onOpen(){
  SpreadsheetApp
  .getUi()
  .createMenu("🎭 CMT")
  .addItem("Apri Portale","apriPortale")
  .addItem("📢 Pubblica aggiornamenti","pubblicaAggiornamenti")
  .addToUi();
}

function apriPortale(){
  const html =
  HtmlService
  .createHtmlOutputFromFile("index")
  .setTitle("🎭 Portale CMT Musical");

  SpreadsheetApp
  .getUi()
  .showSidebar(html);
}

/****************************************************
 FUNZIONI UTILITY DI BASE
****************************************************/

function getFoglio(nome){
  return SpreadsheetApp
  .getActive()
  .getSheetByName(nome);
}

function getDati(nome){
  const sh =
  getFoglio(nome);

  if(!sh){
    return [];
  }

  return sh
  .getDataRange()
  .getValues();
}

function trovaColonna(header,nome){
  for(let i=0;i<header.length;i++){
    if(
      String(header[i])
      .trim()
      .toUpperCase()
      ==
      nome.toUpperCase()
    ){
      return i;
    }
  }
  return -1;
}

// Gestisce "SI", "SÌ", "Sì" in qualsiasi foglio
function valoreSI(v){
  const testo = String(v||"").trim().toUpperCase();
  return testo === "SI" || testo === "SÌ";
}

// Gestisce selezione multipla a tendina (ID CORSI) pulendo spazi e separatori
function listaID(v){
  if(!v){
    return [];
  }
  return String(v)
  .replace(/[\r\n;]+/g, ",") 
  .split(",")
  .map(x => x.trim())
  .filter(x => x !== "");
}

/****************************************************
 GESTIONE VERSIONI GLOBALI CALENDARIO
****************************************************/

function getVersione(){
  let prop = PropertiesService.getScriptProperties();
  let v = prop.getProperty("VERSIONE_CALENDARIO");

  if(!v){
    v = "1";
    prop.setProperty("VERSIONE_CALENDARIO", v);
  }

  return Number(v);
}

function nuovaVersione(){
  let v = getVersione();
  v++;

  PropertiesService
  .getScriptProperties()
  .setProperty("VERSIONE_CALENDARIO", String(v));

  return v;
}

/****************************************************
 ACCESSO PRINCIPALE (FILTRO STAFF / FAMIGLIA)
****************************************************/

function getUtente(email){
  email = String(email).trim().toLowerCase();

  if(!email){
    return {
      errore: true,
      messaggio: "Inserisci una email"
    };
  }

  // 1. Controllo staff/docenti nel foglio ACCESSI
  let staff = cercaAccesso(email);
  if(staff){
    return staff;
  }

  // 2. Controllo accesso famiglie nel foglio GRUPPI
  let famiglia = cercaGruppoEUtenti(email);
  if(famiglia){
    return famiglia;
  }

  return {
    errore: true,
    messaggio: "Email non abilitata o account non attivo nel portale"
  };
}

/****************************************************
 ACCESSI STAFF (FOGLIO ACCESSI)
****************************************************/

function cercaAccesso(email){
  const dati = getDati("ACCESSI");
  if(dati.length < 2) return null;

  const h = dati[0];
  
  const cEmail = trovaColonna(h, "Email docente");
  const cNome = trovaColonna(h, "Nome e Cognome");
  const cRuolo = trovaColonna(h, "Ruolo");
  const cAttivo = trovaColonna(h, "Attivo");
  const cCorsi = trovaColonna(h, "ID CORSI");
  const cTutto = trovaColonna(h, "Vede Tutto");
  const cPub = trovaColonna(h, "Pubblica");
  const cAnteprima = trovaColonna(h, "Anteprima famiglie");

  for(let i=1; i<dati.length; i++){
    if(
      String(dati[i][cEmail]).trim().toLowerCase() == email
      && valoreSI(dati[i][cAttivo])
    ){
      let ruolo = String(dati[i][cRuolo]).trim();
      let èDocente = ruolo.toLowerCase() === "docente";
      let attivaAnteprimaFamiglia = valoreSI(dati[i][cAnteprima]);

      // Se ha l'anteprima famiglia attiva, simula l'esperienza degli allievi
      if (attivaAnteprimaFamiglia) {
        let corsiDocente = listaID(dati[i][cCorsi]);
        return {
          tipo: "famiglia",
          nome: dati[i][cNome] + " (Anteprima Famiglia)",
          ruolo: ruolo,
          corsi: corsiDocente,
          corsiVisualizzati: creaCorsiVisualizzati(corsiDocente),
          vedeTutto: false, 
          pubblica: false,
          puoCaricare: false,
          anteprimaAttiva: true
        };
      }

      // Accesso Staff Standard
      return {
        tipo: "staff",
        nome: dati[i][cNome],
        ruolo: ruolo,
        corsi: listaID(dati[i][cCorsi]),
        vedeTutto: valoreSI(dati[i][cTutto]), 
        pubblica: èDocente ? false : valoreSI(dati[i][cPub]), // I docenti non possono pubblicare modifiche
        puoCaricare: èDocente || ruolo.toLowerCase() === "amministratore" // Flag per abilitare caricamenti Drive
      };
    }
  }
  return null;
}

/****************************************************
 ACCESSI FAMIGLIE (FOGLI GRUPPI + UTENTI)
****************************************************/

function cercaGruppoEUtenti(email){
  const datiGruppi = getDati("GRUPPI");
  if(datiGruppi.length < 2) return null;

  const hGruppi = datiGruppi[0];
  const cGEmail = trovaColonna(hGruppi, "Email d'accesso");
  const cGID = trovaColonna(hGruppi, "ID GRUPPO");
  const cGName = trovaColonna(hGruppi, "Nickname Famiglia");
  const cGAttivo = trovaColonna(hGruppi, "Attivo");

  let idGruppoTrovato = null;
  let nicknameFamiglia = "";

  for(let i=1; i<datiGruppi.length; i++){
    if(
      String(datiGruppi[i][cGEmail]).trim().toLowerCase() == email 
      && valoreSI(datiGruppi[i][cGAttivo])
    ){
      idGruppoTrovato = String(datiGruppi[i][cGID]).trim();
      nicknameFamiglia = String(datiGruppi[i][cGName]).trim();
      break;
    }
  }

  if(!idGruppoTrovato) return null;

  const datiUtenti = getDati("UTENTI");
  if(datiUtenti.length < 2) return null;

  const hUtenti = datiUtenti[0];
  const cUIdGruppo = trovaColonna(hUtenti, "ID GRUPPO");
  const cUAttivo = trovaColonna(hUtenti, "Attivo");

  let persone = [];

  for(let i=1; i<datiUtenti.length; i++){
    if(
      String(datiUtenti[i][cUIdGruppo]).trim() == idGruppoTrovato
      && valoreSI(datiUtenti[i][cUAttivo])
    ){
      persone.push(creaPersona(datiUtenti[i], hUtenti));
    }
  }

  if(persone.length == 0) {
    return {
      errore: true,
      messaggio: "Il gruppo è attivo ma non ci sono allievi associati attivi nel foglio UTENTI."
    };
  }

  // Estraiamo tutti i corsi della famiglia senza doppioni
  let tuttiCorsiID = [];
  persone.forEach(function(p){
    tuttiCorsiID = tuttiCorsiID.concat(p.corsi);
  });
  tuttiCorsiID = [...new Set(tuttiCorsiID)];

  return {
    tipo: "famiglia",
    nome: nicknameFamiglia ? nicknameFamiglia : "Gruppo " + idGruppoTrovato,
    idGruppo: idGruppoTrovato,
    corsi: tuttiCorsiID,
    corsiVisualizzati: creaCorsiVisualizzati(tuttiCorsiID),
    membri: persone
  };
}

function creaPersona(riga, h){
  const cNome = trovaColonna(h, "Nome");
  const cCognome = trovaColonna(h, "Cognome");
  
  let persona = {
    nome: riga[cNome],
    cognome: riga[cCognome],
    corsi: []
  };

  const colInizioCorsi = trovaColonna(h, "ID GRUPPO") + 1;

  for(let i = colInizioCorsi; i < h.length; i++){
    let codiceCorso = String(h[i]).trim();
    let valorecella = String(riga[i]).trim().toUpperCase();

    if(valorecella == "X"){
      persona.corsi.push(codiceCorso);
    }
  }

  return persona;
}

/****************************************************
 CONVERSIONE ID CORSO -> NOME CORSO (FOGLIO CORSI)
****************************************************/

function nomeCorso(id){
  const dati = getDati("CORSI");
  if(dati.length < 2) return id;

  const h = dati[0];
  const cID = trovaColonna(h, "ID CORSI");
  const cNome = trovaColonna(h, "NOME CORSO");

  if (cID === -1 || cNome === -1) return id;

  const idCercato = String(id).trim().toUpperCase();

  for(let i = 1; i < dati.length; i++){
    if(String(dati[i][cID]).trim().toUpperCase() === idCercato){
      return dati[i][cNome];
    }
  }
  return id;
}

function creaCorsiVisualizzati(corsi){
  let risultato = [];
  corsi.forEach(function(id){
    risultato.push({
      id: id,
      nome: nomeCorso(id)
    });
  });
  return risultato;
}

/****************************************************
 CALENDARIO PROVE (CON STATO BOZZA/PUBBLICATO)
****************************************************/

function getCalendario(corsi, vedeTutto, tipoUtente){
  const dati = getDati("CALENDARIO");
  if(dati.length < 2) return [];

  const h = dati[0];
  const cData = trovaColonna(h, "Data");
  const cOra = trovaColonna(h, "Orario");
  const cCosa = trovaColonna(h, "Cosa si fa");
  const cDest = trovaColonna(h, "Destinatari");
  const cStato = trovaColonna(h, "Stato");
  const cNote = trovaColonna(h, "Note Pubbliche");
  const cVersione = trovaColonna(h, "Versione");

  let risultati = [];

  for(let i=1;i<dati.length;i++){
    let stato = String(dati[i][cStato]).trim().toUpperCase();

    // Le famiglie vedono SOLO i turni PUBBLICATO. Lo staff vede tutto.
    if(tipoUtente === "famiglia" && stato !== "PUBBLICATO"){
      continue;
    }

    let destinatari = String(dati[i][cDest])
    .split(",")
    .map(x=>x.trim())
    .filter(x=>x!="");

    let trovato = false;

    if(vedeTutto){
      trovato = true;
    } else {
      corsi.forEach(function(c){
        if(destinatari.includes(c)){
          trovato = true;
        }
      });
    }

    if(!trovato) continue;

    risultati.push({
      data: formattaData(dati[i][cData]),
      orario: dati[i][cOra],
      cosa: dati[i][cCosa],
      destinatari: destinatari,
      stato: stato,
      note: dati[i][cNote],
      versione: Number(dati[i][cVersione] || 1)
    });
  }
  return risultati;
}

function formattaData(data){
  if(Object.prototype.toString.call(data) == "[object Date]"){
    return Utilities.formatDate(data, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return data;
}

/****************************************************
 NOTIFICA LETTURE CALENDARIO (FOGLIO VERSIONI)
****************************************************/

function getVersioneVista(email){
  const sh = getFoglio("VERSIONI");
  if(!sh) return 0;

  const dati = sh.getDataRange().getValues();
  if(dati.length < 2) return 0;

  const h = dati[0];
  const cEmail = trovaColonna(h, "Email");
  const cVersioneVista = trovaColonna(h, "Versione Vista");

  if (cEmail === -1 || cVersioneVista === -1) return 0;

  for(let i=1; i<dati.length; i++){
    if(String(dati[i][cEmail]).trim().toLowerCase() == email.toLowerCase()){
      return Number(dati[i][cVersioneVista] || 0);
    }
  }
  return 0;
}

function salvaVersioneVista(email, versione){
  const sh = getFoglio("VERSIONI");
  if(!sh) return;

  const dati = sh.getDataRange().getValues();
  const h = dati[0];
  const cEmail = trovaColonna(h, "Email");
  const cVersioneVista = trovaColonna(h, "Versione Vista");

  if (cEmail === -1 || cVersioneVista === -1) {
    sh.clear();
    sh.appendRow(["Email", "Versione Vista"]);
    sh.appendRow([email, versione]);
    return;
  }

  for(let i=1; i<dati.length; i++){
    if(String(dati[i][cEmail]).trim().toLowerCase() == email.toLowerCase()){
      sh.getRange(i+1, cVersioneVista+1).setValue(versione);
      return;
    }
  }

  let nuovaRiga = [];
  nuovaRiga[cEmail] = email;
  nuovaRiga[cVersioneVista] = versione;
  
  sh.appendRow(nuovaRiga);
}

/****************************************************
 PUBBLICAZIONE CALENDARIO
****************************************************/

function pubblicaAggiornamenti(){
  let versione = nuovaVersione();
  const sh = getFoglio("CALENDARIO");
  const dati = sh.getDataRange().getValues();
  const h = dati[0];

  const cStato = trovaColonna(h, "Stato");
  const cVersione = trovaColonna(h, "Versione");

  for(let i=1;i<dati.length;i++){
    let stato = String(dati[i][cStato]).trim().toUpperCase();
    if(stato == "PUBBLICATO"){
      sh.getRange(i+1, cVersione+1).setValue(versione);
    }
  }

  PropertiesService
  .getScriptProperties()
  .setProperty("ULTIMA_PUBBLICAZIONE", new Date().toString());

  SpreadsheetApp
  .getUi()
  .alert("📢 Aggiornamento pubblicato\nVersione " + versione);
}

/****************************************************
 MATERIALI DIDATTICI (LETTURA E SCRITTURA)
****************************************************/

function getMateriali(corsi){
  const dati = getDati("MATERIALI");
  if(dati.length < 2) return [];

  const h = dati[0];
  const cID = trovaColonna(h, "ID CORSI");
  const cTitolo = trovaColonna(h, "Titolo");
  const cDesc = trovaColonna(h, "Descrizione");
  const cLink = trovaColonna(h, "Link Drive");
  const cPubblicato = trovaColonna(h, "Pubblicato");

  if (cID === -1 || cTitolo === -1 || cLink === -1) return [];

  let risultato = [];
  for(let i=1; i<dati.length; i++){
    let idCorsoRiga = String(dati[i][cID]).trim();
    let isPubblicato = valoreSI(dati[i][cPubblicato]);

    if(corsi.includes(idCorsoRiga) && isPubblicato){
      risultato.push({
        corso: nomeCorso(idCorsoRiga),
        titolo: dati[i][cTitolo],
        descrizione: cDesc !== -1 ? dati[i][cDesc] : "",
        link: dati[i][cLink]
      });
    }
  }
  return risultato;
}

function salvaNuovoMateriale(idCorso, titolo, descrizione, linkDrive) {
  const sh = getFoglio("MATERIALI");
  if(!sh) return { errore: true, messaggio: "Foglio MATERIALI non trovato." };
  
  const dati = sh.getDataRange().getValues();
  const h = dati[0];

  const cID = trovaColonna(h, "ID CORSI");
  const cTitolo = trovaColonna(h, "Titolo");
  const cDesc = trovaColonna(h, "Descrizione");
  const cLink = trovaColonna(h, "Link Drive");
  const cPubblicato = trovaColonna(h, "Pubblicato");
  const cData = trovaColonna(h, "Data");

  if (cID === -1 || cTitolo === -1 || cLink === -1) {
    return { errore: true, messaggio: "Struttura delle colonne del foglio MATERIALI non valida." };
  }

  let nuovaRiga = [];
  nuovaRiga[cID] = idCorso;
  nuovaRiga[cTitolo] = titolo;
  if (cDesc !== -1) nuovaRiga[cDesc] = descrizione;
  nuovaRiga[cLink] = linkDrive;
  
  if (cPubblicato !== -1) nuovaRiga[cPubblicato] = "SI";
  if (cData !== -1) {
    nuovaRiga[cData] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  
  sh.appendRow(nuovaRiga);
  return { errore: false, messaggio: "Materiale caricato e salvato correttamente!" };
}

/****************************************************
 CARICAMENTO GLOBALE PORTALE
****************************************************/

function controlloAggiornamenti(email){
  return (getVersioneVista(email) < getVersione());
}

function getUltimaPubblicazione(){
  let d = PropertiesService.getScriptProperties().getProperty("ULTIMA_PUBBLICAZIONE");
  if(!d) return "";
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
}

function caricaPortale(email){
  let utente = getUtente(email);
  if(utente.errore) return utente;

  let corsi = utente.corsi || [];
  let calendario = getCalendario(corsi, utente.vedeTutto || false, utente.tipo);

  return {
    utente: utente,
    corsiVisualizzati: creaCorsiVisualizzati(corsi),
    prove: calendario,
    materiali: getMateriali(corsi),
    versione: getVersione(),
    versioneVista: getVersioneVista(email),
    aggiornamenti: utente.tipo === "famiglia" ? controlloAggiornamenti(email) : false,
    ultimaPubblicazione: getUltimaPubblicazione()
  };
}

function confermaVisione(email){
  salvaVersioneVista(email, getVersione());
  return true;
} 