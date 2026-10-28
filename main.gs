/****************************************************
 🎭 CMT MUSICAL - PORTALE COMPLETO (SISTEMA UNIFICATO)
 ****************************************************/


/**
 * Servizio Web App: rende il portale accessibile via URL del browser
 */
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("🎭 Portale CMT Musical")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
/**
 * Crea il menu personalizzato nella barra del foglio Google Sheets
 * Viene eseguita automaticamente all'apertura del foglio
 * Menu items:
 * - Apri Portale: visualizza la sidebar HTML
 * - Pubblica aggiornamenti: incrementa versione e notifica famiglie
 */
function onOpen(){
  SpreadsheetApp
  .getUi()
  .createMenu("🎭 CMT")
  .addItem("Apri Portale","apriPortale")
  .addItem("📢 Pubblica aggiornamenti","pubblicaAggiornamenti")
  .addToUi();
}

/**
 * Apre il portale HTML come sidebar nel foglio Sheets
 * Carica il file index.html e lo mostra accanto al foglio
 */
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
 FUNZIONI UTILITY DI BASE - Accesso dati generico
****************************************************/

const cacheDati = {};
let cacheNomiCorsi = null;

/**
 * Recupera un foglio per nome
 * @param {string} nome - Nome del foglio (es. "CALENDARIO", "ACCESSI")
 * @returns {Sheet} Oggetto foglio, null se non trovato
 */
function getFoglio(nome){
  return SpreadsheetApp
  .getActive()
  .getSheetByName(nome);
}

/**
 * Legge tutti i dati di un foglio
 * Utile per accedere alla struttura completa (header + righe)
 * @param {string} nome - Nome del foglio
 * @returns {Array<Array>} Matrice 2D [righe][colonne]. Riga 0 è l'header
 */
function getDati(nome){
  if(Object.prototype.hasOwnProperty.call(cacheDati, nome)){
    return cacheDati[nome];
  }

  const sh = getFoglio(nome);

  if(!sh){
    cacheDati[nome] = [];
    return [];
  }

  cacheDati[nome] = sh.getDataRange().getValues();

  return cacheDati[nome];
}

function getFoglioCorsi(){
  const chiaveCache = "CORSI";
  return SpreadsheetApp
    .getActive()
    .getSheets()
    .find(function(foglio){
      return String(foglio.getName()).trim().toUpperCase() === chiaveCache;
    }) || null;
}

function getDatiCorsi(){
  const chiaveCache = "CORSI";
  if(Object.prototype.hasOwnProperty.call(cacheDati, chiaveCache)){
    return cacheDati[chiaveCache];
  }

  const foglioCorsi = getFoglioCorsi();

  cacheDati[chiaveCache] = foglioCorsi
    ? foglioCorsi.getDataRange().getDisplayValues()
    : [];
  return cacheDati[chiaveCache];
}

/**
 * Trova l'indice di una colonna basato sul nome nell'header
 * Case-insensitive, tollera spazi extra
 * @param {Array} header - Prima riga (intestazioni colonne)
 * @param {string} nome - Nome colonna da cercare
 * @returns {number} Indice 0-based della colonna, -1 se non trovata
 */
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
 Sistema di notifiche per aggiornamenti
****************************************************/

/**
 * Ottiene il numero versione attuale del calendario
 * Utilizzato per notificare le famiglie di aggiornamenti
 * Inizializza a 1 se non esiste
 * Memorizzato in PropertiesService (persistente tra sessioni)
 * @returns {number} Versione attuale del calendario
 */
function getVersione(){
  let prop = PropertiesService.getScriptProperties();
  let v = prop.getProperty("VERSIONE_CALENDARIO");

  if(!v){
    v = "1";
    prop.setProperty("VERSIONE_CALENDARIO", v);
  }

  return Number(v);
}

/**
 * Incrementa la versione del calendario di 1
 * Chiamato quando si clicca "📢 Pubblica aggiornamenti"
 * Le famiglie vedranno una notifica al prossimo accesso
 * @returns {number} Nuovo numero versione
 */
function nuovaVersione(){
  let v = getVersione();
  v++;

  PropertiesService
  .getScriptProperties()
  .setProperty("VERSIONE_CALENDARIO", String(v));

  return v;
}

/****************************************************
 ACCESSO PRINCIPALE - Discrimina STAFF vs FAMIGLIA
 Entry point di autenticazione
****************************************************/

/**
 * FUNZIONE PRINCIPALE DI AUTENTICAZIONE
 * Determina il tipo di utente e carica i dati appropriati
 * Flusso: Staff (ACCESSI) → Gruppo/Individuale (GRUPPI+UTENTI) → Errore
 * 
 * @param {string} email - Email dell'utente che accede
 * @returns {Object} 
 *   - Successo: {tipo, nome, ruolo, corsi, vedeTutto, ...}
 *   - Errore: {errore: true, messaggio: "..."}  
 */
function getUtente(email){
  email = String(email).trim().toLowerCase();

  if(!email){
    return {
      errore: true,
      messaggio: "Inserisci una email"
    };
  }

  // STEP 1: Prova a trovare l'utente come STAFF/DOCENTE (foglio ACCESSI)
  let staff = cercaAccesso(email);
  if(staff){
    return staff;
  }

  // STEP 2: Se non staff, prova FAMIGLIA (foglio GRUPPI + UTENTI)
  let famiglia = cercaGruppoEUtenti(email);
  if(famiglia){
    return famiglia;
  }

  // STEP 3: Email non riconosciuta
  return {
    errore: true,
    messaggio: "Email non abilitata o account non attivo nel portale"
  };
}

/****************************************************
 ACCESSI STAFF - Foglio "ACCESSI"
 Docenti, Amministratori, e staff con ruoli specifici
****************************************************/

/**
 * Cerca un utente nel foglio ACCESSI (Staff/Docenti)
 * Se trovato E ATTIVO, ritorna il profilo con permessi specifici
 * 
 * Colonne attese nel foglio ACCESSI:
 * - Email docente (univoca)
 * - Nome e Cognome
 * - Ruolo ("docente", "amministratore", etc.)
 * - Attivo ("SI"/"SÌ" per abilitare)
 * - ID CORSI (corsi gestiti da questo staff)
 * - Vede Tutto ("SI" = vede tutti i turni)
 * - Pubblica ("SI" = può modificare stato turni)
 * - Anteprima famiglie ("SI" = simula esperienza allievato)
 * 
 * @param {string} email - Email dello staff
 * @returns {Object|null} Profilo staff se trovato, null altrimenti
 */
function cercaAccesso(email){
  const dati = getDati("ACCESSI");
  if(dati.length < 2) return null;  // Foglio vuoto o solo header

  const h = dati[0];
  
  // Mappa colonne con i nomi attesi
  const cEmail = trovaColonna(h, "Email docente");
  const cNome = trovaColonna(h, "Nome e Cognome");
  const cRuolo = trovaColonna(h, "Ruolo");
  const cAttivo = trovaColonna(h, "Attivo");
  const cCorsi = trovaColonna(h, "ID CORSI");
  const cTutto = trovaColonna(h, "Vede Tutto");
  const cPub = trovaColonna(h, "Pubblica");
  const cAnteprima = trovaColonna(h, "Anteprima famiglie");

  // Itera su tutte le righe di staff (partendo da 1, saltando header)
  for(let i=1; i<dati.length; i++){
    // Verifica: email corrisponde E account è ATTIVO
    if(
      String(dati[i][cEmail]).trim().toLowerCase() == email
      && valoreSI(dati[i][cAttivo])
    ){
      let ruolo = String(dati[i][cRuolo]).trim();
      let èDocente = ruolo.toLowerCase() === "docente";
      let attivaAnteprimaFamiglia = valoreSI(dati[i][cAnteprima]);

      // CASO SPECIALE: Anteprima Famiglia
      // Permette ai docenti di testare l'esperienza dell'allievato
      if (attivaAnteprimaFamiglia) {
        let corsiDocente = listaID(dati[i][cCorsi]);
        return {
          tipo: "famiglia",  // Simula esperienza famiglia
          nome: dati[i][cNome] + " (Anteprima Famiglia)",
          ruolo: ruolo,
          corsi: corsiDocente,
          corsiVisualizzati: creaCorsiVisualizzati(corsiDocente),
          vedeTutto: false,       // Non vede tutto
          pubblica: false,        // Non può pubblicare
          puoCaricare: false,     // Non può caricare
          anteprimaAttiva: true   // Flag per UI
        };
      }

      // ACCESSO STAFF STANDARD
      return {
        tipo: "staff",
        nome: dati[i][cNome],
        ruolo: ruolo,
        corsi: listaID(dati[i][cCorsi]),
        vedeTutto: valoreSI(dati[i][cTutto]),  // Può vedere tutti i turni
        // IMPORTANTE: Docenti NON possono pubblicare modifiche (limitazione di sicurezza)
        pubblica: èDocente ? false : valoreSI(dati[i][cPub]),
        // Abilitazione caricamento file Drive (docenti + admin)
        puoCaricare: èDocente || ruolo.toLowerCase() === "amministratore"
      };
    }
  }
  return null;  // Email non trovata in ACCESSI
}

/****************************************************
 ACCESSI FAMIGLIE E INDIVIDUALI - Fogli "GRUPPI" + "UTENTI"
 Schema: Un account (email) → ID GRUPPO opzionale → Uno o più utenti
****************************************************/

/**
 * Cerca un utente nel sistema di famiglie (GRUPPI + UTENTI)
 * 
 * Processo in 2 step:
 * 1. Trova ID GRUPPO dalla email nel foglio GRUPPI
 * 2. Se l'ID è vuoto, legge l'utente individuale senza gruppo
 * 3. Altrimenti legge tutti gli allievi associati al gruppo da UTENTI
 * 
 * Colonne attese in GRUPPI:
 * - Email d'accesso
 * - ID GRUPPO (univoco)
 * - Nickname Famiglia (nome display)
 * - Attivo ("SI"/"SÌ")
 * 
 * Colonne attese in UTENTI:
 * - ID GRUPPO (link a GRUPPI)
 * - Nome, Cognome
 * - Attivo ("SI"/"SÌ")
 * - Colonne da ID GRUPPO+1 in poi: codici corsi (marcati con "X")
 * 
 * @param {string} email - Email del genitore
 * @returns {Object|null} Profilo famiglia/individuale, null se non trovata
 */
function cercaGruppoEUtenti(email){
  const datiGruppi = getDati("GRUPPI");
  if(datiGruppi.length < 2) return null;  // Foglio GRUPPI vuoto

  const hGruppi = datiGruppi[0];
  // Mappa colonne GRUPPI
  const cGEmail = trovaColonna(hGruppi, "Email d'accesso");
  const cGID = trovaColonna(hGruppi, "ID GRUPPO");
  const cGName = trovaColonna(hGruppi, "Nickname Famiglia");
  const cGAttivo = trovaColonna(hGruppi, "Attivo");

  let idGruppoTrovato = null;
  let nicknameFamiglia = "";

  // STEP 1: Trova il gruppo della famiglia da email
  if(cGEmail !== -1 && cGID !== -1){
    for(let i=1; i<datiGruppi.length; i++){
      let idGruppo = String(datiGruppi[i][cGID] || "").trim();
      if(
        String(datiGruppi[i][cGEmail]).trim().toLowerCase() == email
        && idGruppo
        && valoreSI(datiGruppi[i][cGAttivo])  // Account genitore deve essere attivo
      ){
        idGruppoTrovato = idGruppo;
        nicknameFamiglia = cGName === -1 ? "" : String(datiGruppi[i][cGName]).trim();
        break;
      }
    }
  }

  // STEP 2: Leggi l'utente individuale o gli allievi associati al gruppo
  const datiUtenti = getDati("UTENTI");
  if(datiUtenti.length < 2) return null;  // Foglio UTENTI vuoto

  const hUtenti = datiUtenti[0];
  // Mappa colonne UTENTI
  const cUIdGruppo = trovaColonna(hUtenti, "ID GRUPPO");
  const cUAttivo = trovaColonna(hUtenti, "Attivo");
  const cUEmail = ["Email", "Email d'accesso", "Email utente"]
    .map(function(nome) { return trovaColonna(hUtenti, nome); })
    .find(function(indice) { return indice !== -1; });

  let persone = [];  // Array di allievi della famiglia
  let accountIndividuale = !idGruppoTrovato;

  // Raccogli l'utente individuale o tutti gli allievi attivi del gruppo
  for(let i=1; i<datiUtenti.length; i++){
    let attivo = valoreSI(datiUtenti[i][cUAttivo]);
    let idRiga = cUIdGruppo === -1 ? "" : String(datiUtenti[i][cUIdGruppo]).trim();
    let emailRiga = cUEmail === undefined ? "" : String(datiUtenti[i][cUEmail]).trim().toLowerCase();
    let corrisponde = idGruppoTrovato
      ? idRiga == idGruppoTrovato
      : (cUEmail !== undefined ? emailRiga == email : idRiga == "");

    if(attivo && corrisponde){
      persone.push(creaPersona(datiUtenti[i], hUtenti));
      if(accountIndividuale && cUEmail === undefined) break;
    }
  }

  // Controllo: il gruppo esiste ma senza allievi attivi
  if(persone.length == 0) {
    return {
      errore: true,
      messaggio: "Il gruppo è attivo ma non ci sono allievi associati attivi nel foglio UTENTI."
    };
  }

  // Unione di tutti i corsi dei vari allievi (senza duplicati)
  // Serve per mostrare il calendario rilevante per la famiglia
  let tuttiCorsiID = [];
  persone.forEach(function(p){
    tuttiCorsiID = tuttiCorsiID.concat(p.corsi);
  });
  tuttiCorsiID = [...new Set(tuttiCorsiID)];  // Rimuovi duplicati

  // Profilo individuale o famiglia con lista completa di allievi
  return {
    tipo: accountIndividuale ? "individuale" : "famiglia",
    nome: accountIndividuale
      ? persone[0].nome + " " + persone[0].cognome
      : (nicknameFamiglia ? nicknameFamiglia : "Gruppo " + idGruppoTrovato),
    idGruppo: idGruppoTrovato,
    famigliaVerificata: !accountIndividuale && Boolean(idGruppoTrovato),
    corsi: tuttiCorsiID,  // Elenco unico di corsi per la famiglia
    corsiVisualizzati: creaCorsiVisualizzati(tuttiCorsiID),
    membri: persone  // Lista allievi (con loro corsi singoli)
  };
}

/**
 * Estrae dati di un singolo allievato da una riga del foglio UTENTI
 * Le colonne corso partono dopo "ID GRUPPO"
 * Un allievato è iscritto a un corso se la cella corrispond contiene "X"
 * 
 * @param {Array} riga - Una riga del foglio UTENTI
 * @param {Array} h - Header del foglio UTENTI
 * @returns {Object} {nome, cognome, corsi: [id1, id2, ...]}
 */
function creaPersona(riga, h){
  const cNome = trovaColonna(h, "Nome");
  const cCognome = trovaColonna(h, "Cognome");
  
  let persona = {
    nome: riga[cNome],
    cognome: riga[cCognome],
    corsi: []  // Elenco corsi iscritti
  };

  // Le colonne corso cominciano dopo "ID GRUPPO"
  // Es: col 0="ID GRUPPO", col 1="Nome", col 2="Cognome", col 3+="PIANO", "CANTO", etc.
  const colInizioCorsi = trovaColonna(h, "ID GRUPPO") + 1;

  // Itera sulle colonne corso e raccogli quelle marcate con "X"
  for(let i = colInizioCorsi; i < h.length; i++){
    let codiceCorso = String(h[i]).trim();  // Nome del corso dalla header
    let valorecella = String(riga[i]).trim().toUpperCase();

    // Se marcato con "X", l'allievato è iscritto a questo corso
    if(valorecella == "X"){
      persona.corsi.push(codiceCorso);
    }
  }

  return persona;
}

/****************************************************
 LOOKUP CORSO - Conversione ID → NOME CORSO
 Foglio CORSI contiene la mappatura
****************************************************/

/**
 * Converte un ID corso nel suo nome descrittivo
 * Es: "PIANO" → "Lezioni di Pianoforte A.A. 2025/26"
 * Se non trovato, ritorna l'ID originale
 * 
 * @param {string} id - ID corso (codice breve)
 * @returns {string} Nome completo del corso
 */
function nomeCorso(id){
  const idCercato = String(id).trim().toUpperCase();
  const mappa = getMappaNomiCorsi();
  return Object.prototype.hasOwnProperty.call(mappa, idCercato)
    ? mappa[idCercato]
    : id;
}

function getMappaNomiCorsi(){
  if(cacheNomiCorsi !== null) return cacheNomiCorsi;

  cacheNomiCorsi = {};
  const dati = getDatiCorsi();
  if(dati.length < 2) return cacheNomiCorsi;

  const h = dati[0];
  const cID = trovaColonna(h, "ID CORSI");
  const cNomeDrive = trovaColonna(h, "nome corso e cartella drive");
  const cNome = cNomeDrive !== -1 ? cNomeDrive : trovaColonna(h, "NOME CORSO");
  if(cID === -1 || cNome === -1) return cacheNomiCorsi;

  for(let i = 1; i < dati.length; i++){
    const idCorso = String(dati[i][cID]).trim().toUpperCase();
    if(idCorso) cacheNomiCorsi[idCorso] = dati[i][cNome];
  }
  return cacheNomiCorsi;
}

function trovaColonnaFlessibile(header, alternative){
  const intestazioni = header.map(function(valore){
    return normalizzaCorso(valore);
  });

  for(let i = 0; i < alternative.length; i++){
    const nome = normalizzaCorso(alternative[i]);
    const indiceEsatto = intestazioni.indexOf(nome);
    if(indiceEsatto !== -1) return indiceEsatto;
  }

  for(let i = 0; i < alternative.length; i++){
    const nome = normalizzaCorso(alternative[i]);
    const indiceParziale = intestazioni.findIndex(function(intestazione){
      return nome.length > 3 && intestazione.indexOf(nome) !== -1;
    });
    if(indiceParziale !== -1) return indiceParziale;
  }
  return -1;
}

function parseOrariLezioni(testoOrari, giornoAlternativo){
  const giorni = [
    {nome: "Lunedì", pattern: "lunedi"},
    {nome: "Martedì", pattern: "martedi"},
    {nome: "Mercoledì", pattern: "mercoledi"},
    {nome: "Giovedì", pattern: "giovedi"},
    {nome: "Venerdì", pattern: "venerdi"},
    {nome: "Sabato", pattern: "sabato"},
    {nome: "Domenica", pattern: "domenica"}
  ];
  const occasioni = [];
  const parti = String(testoOrari || "").split(/[\r\n;|]+/);

  parti.forEach(function(parte){
    const testo = String(parte || "").normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();
    const regexGiorno = /\b(lunedi|lun|martedi|mar|mercoledi|mer|giovedi|gio|venerdi|ven|sabato|sab|domenica|dom)\b/gi;
    const giorniTrovati = [];
    let corrispondenzaGiorno;
    while((corrispondenzaGiorno = regexGiorno.exec(testo)) !== null){
      const giornoNormalizzato = normalizzaCorso(corrispondenzaGiorno[1]);
      const giorno = giorni.find(function(candidato){
        return candidato.pattern.toUpperCase().indexOf(giornoNormalizzato) === 0;
      });
      if(giorno && giorniTrovati.indexOf(giorno) === -1){
        giorniTrovati.push(giorno);
      }
    }
    if(giorniTrovati.length === 0 && giornoAlternativo){
      const giornoNormalizzato = normalizzaCorso(giornoAlternativo);
      const giorno = giorni.find(function(candidato){
        const pattern = candidato.pattern.toUpperCase();
        return pattern.indexOf(giornoNormalizzato) === 0 ||
          giornoNormalizzato.indexOf(pattern) === 0;
      });
      if(giorno) giorniTrovati.push(giorno);
    }

    const intervalli = [];
    const regexOrario = /(\d{1,2})\s*[:.]\s*(\d{2})\s*(?:-|–|—|\/|\b(?:a|alle)\b|\bfino\s+alle\b)\s*(\d{1,2})\s*[:.]\s*(\d{2})/gi;
    let corrispondenza;
    while((corrispondenza = regexOrario.exec(parte)) !== null){
      const inizioOra = Number(corrispondenza[1]);
      const inizioMinuti = Number(corrispondenza[2]);
      const fineOra = Number(corrispondenza[3]);
      const fineMinuti = Number(corrispondenza[4]);
      const inizio = inizioOra * 60 + inizioMinuti;
      const fine = fineOra * 60 + fineMinuti;
      if(inizioOra < 24 && fineOra < 24 && inizioMinuti < 60 &&
        fineMinuti < 60 && fine > inizio){
        intervalli.push({
          inizio: String(inizioOra).padStart(2, "0") + ":" + String(inizioMinuti).padStart(2, "0"),
          fine: String(fineOra).padStart(2, "0") + ":" + String(fineMinuti).padStart(2, "0"),
          minutiInizio: inizio,
          minutiFine: fine
        });
      }
    }

    giorniTrovati.forEach(function(giorno){
      intervalli.forEach(function(intervallo){
        occasioni.push({
          giorno: giorno.nome,
          inizio: intervallo.inizio,
          fine: intervallo.fine,
          minutiInizio: intervallo.minutiInizio,
          minutiFine: intervallo.minutiFine
        });
      });
    });
  });

  return occasioni;
}

function getOrariLezioni(corsi){
  const foglioCorsi = getFoglioCorsi();
  if(!foglioCorsi || foglioCorsi.getLastRow() < 2) return [];

  const datiCompleti = foglioCorsi
    .getRange(1, 1, foglioCorsi.getLastRow(), 9)
    .getDisplayValues();
  const headerCompleto = datiCompleti[0];
  const dati = foglioCorsi
    .getRange(1, 4, foglioCorsi.getLastRow(), 6)
    .getDisplayValues();
  const h = dati[0];
  const cID = trovaColonnaFlessibile(headerCompleto, [
    "ID CORSI", "ID CORSO", "CODICE CORSO", "ID"
  ]);
  const cMateriaHeader = trovaColonnaFlessibile(h, [
    "MATERIA", "NOME CORSO E CARTELLA DRIVE", "NOME CORSO", "CORSO"
  ]);
  const cEtaHeader = trovaColonnaFlessibile(h, [
    "ETA", "FASCIA ETA", "FASCIA D ETA", "FASCIA DI ETA", "ETA ALLIEVI", "ANNI"
  ]);
  const cDocenteHeader = trovaColonnaFlessibile(h, [
    "DOCENTE", "DOCENTI", "INSEGNANTE", "NOME DOCENTE"
  ]);
  const cOrariHeader = trovaColonnaFlessibile(h, ["ORARI", "ORARIO"]);
  const cGiorno = trovaColonnaFlessibile(h, ["GIORNO", "GIORNI"]);
  const cMateria = cMateriaHeader !== -1 ? cMateriaHeader : 0;
  const cEta = cEtaHeader !== -1 ? cEtaHeader : 1;
  const cDocente = cDocenteHeader !== -1 ? cDocenteHeader : 2;
  const cOrari = cOrariHeader !== -1 ? cOrariHeader : 5;
  const coloriCorsi = foglioCorsi
    .getRange(2, 4 + cMateria, dati.length - 1, 1)
    .getBackgrounds();

  const corsiConsentiti = Object.create(null);
  (corsi || []).forEach(function(corso){
    const id = normalizzaCorso(corso);
    const nome = normalizzaCorso(nomeCorso(corso));
    if(id) corsiConsentiti[id] = true;
    if(nome) corsiConsentiti[nome] = true;
  });
  let risultati = [];

  for(let i = 1; i < dati.length; i++){
    const rigaCorso = dati[i];
    const materia = String(rigaCorso[cMateria] || "").trim();
    const idCorso = cID === -1 ? "" : String(datiCompleti[i][cID] || "").trim();
    const idNormalizzato = normalizzaCorso(idCorso);
    const materiaNormalizzata = normalizzaCorso(materia);
    const nomeDaID = normalizzaCorso(idCorso ? nomeCorso(idCorso) : "");
    const isCompagnia = materiaNormalizzata === "COMPAGNIA";
    const corrisponde = isCompagnia
      ? Boolean(corsiConsentiti["COMPAGNIA2627"])
      : Boolean(
        corsiConsentiti[idNormalizzato] ||
        corsiConsentiti[materiaNormalizzata] ||
        corsiConsentiti[nomeDaID]
      );
    if(!corrisponde) continue;

    const testoOrari = rigaCorso[cOrari];
    const giornoAlternativo = isCompagnia
      ? "Mercoledì"
      : cGiorno === -1 ? "" : rigaCorso[cGiorno];
    const occasioni = parseOrariLezioni(testoOrari, giornoAlternativo);
    if(isCompagnia){
      occasioni.forEach(function(occasione){
        occasione.giorno = "Mercoledì";
      });
    }
    occasioni.forEach(function(occasione){
      risultati.push({
        id: idCorso || materia,
        materia: materia || "Corso",
        eta: String(rigaCorso[cEta] || "").trim(),
        docente: String(rigaCorso[cDocente] || "").trim(),
        giorno: occasione.giorno,
        inizio: occasione.inizio,
        fine: occasione.fine,
        minutiInizio: occasione.minutiInizio,
        minutiFine: occasione.minutiFine,
        colore: coloriCorsi.length > 0 ? coloriCorsi[i - 1][0] : ""
      });
    });
  }

  return risultati;
}

function normalizzaCorso(valore){
  return String(valore || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "")
    .trim();
}

/**
 * Verifica se un destinatario identifica uno dei corsi dell'utente.
 * Accetta sia l'ID tecnico sia il nome del corso nella scheda CORSI.
 */
function destinatarioCorrisponde(destinatario, corsi){
  const valore = normalizzaCorso(destinatario);
  if(!valore) return false;

  return corsi.some(function(corso){
    const id = normalizzaCorso(corso);
    const nome = normalizzaCorso(nomeCorso(corso));
    return valore === id || valore === nome ||
      (nome.length > 3 && (valore.includes(nome) || nome.includes(valore)));
  });
}

/**
 * Converte array di ID corsi in array di oggetti {id, nome}
 * Usato per popolare UI con nomi leggibili
 * 
 * @param {Array<string>} corsi - Array di ID corsi
 * @returns {Array<Object>} Array di {id, nome}
 */
function creaCorsiVisualizzati(corsi){
  let risultato = [];
  corsi.forEach(function(id){
    risultato.push({
      id: id,
      nome: nomeCorso(id)  // Lookup del nome completo
    });
  });
  return risultato;
}

/****************************************************
 CALENDARIO PROVE - Filtro per tipo utente
 Logica di accesso: Staff vede tutto | Famiglia vede PUBBLICATO
****************************************************/

/**
 * Recupera le prove del calendario filtrando per:
 * 1. Tipo utente (staff vs famiglia)
 * 2. Stato turno (BOZZA vs PUBBLICATO)
 * 3. Corsi assegnati all'utente
 * 
 * REGOLA CRITICA: 
 * - Staff vede TUTTO (BOZZA + PUBBLICATO)
 * - Famiglia vede SOLO PUBBLICATO
 * 
 * @param {Array<string>} corsi - Corsi assegnati all'utente
 * @param {boolean} vedeTutto - Se true, ignora filtri corsi (staff admin)
 * @param {string} tipoUtente - "staff" o "famiglia"
 * Le prime quattro colonne del foglio sono: Data, Orario, Prova, Destinatari.
 * Le colonne Stato, Note Pubbliche e Versione sono opzionali.
 * @returns {Array<Object>} Prove filtrate [{data, orario, cosa, colore, destinatari, stato, note, versione}]
 */
function getCalendario(corsi, vedeTutto, tipoUtente){
  const dati = getDati("CALENDARIO");
  if(dati.length < 2) return [];  // Nessun calendario

  const h = dati[0];
  // A-D sono il formato minimo per ogni prova; i nomi delle intestazioni
  // vengono preferiti, con fallback alle prime quattro colonne.
  const cData = trovaColonna(h, "Data") !== -1 ? trovaColonna(h, "Data") : 0;
  const cOra = trovaColonna(h, "Orario") !== -1 ? trovaColonna(h, "Orario") : 1;
  const cCosa = trovaColonna(h, "Cosa si fa") !== -1 ? trovaColonna(h, "Cosa si fa") : 2;
  const cDest = trovaColonna(h, "Destinatari") !== -1 ? trovaColonna(h, "Destinatari") : 3;
  const cStato = trovaColonna(h, "Stato");
  const cNote = trovaColonna(h, "Note Pubbliche");
  const cVersione = trovaColonna(h, "Versione");
  const foglio = getFoglio("CALENDARIO");
  const coloriAttivita = foglio
    .getRange(2, cCosa + 1, dati.length - 1, 1)
    .getBackgrounds();

  let risultati = [];

  // Itera su tutti i turni
  for(let i=1;i<dati.length;i++){
    let statoLetto = cStato === -1
      ? "PUBBLICATO"
      : String(dati[i][cStato]).trim().toUpperCase();
    let stato = (!statoLetto || valoreSI(statoLetto) || statoLetto === "PUBBLICATO")
      ? "PUBBLICATO"
      : statoLetto;

    // FILTRO ACCESSO: Famiglie vedono solo PUBBLICATO, staff vede tutto
    // Questo è un controllo di sicurezza critico
    if(tipoUtente !== "staff" && stato !== "PUBBLICATO"){
      continue;  // Nasconde BOZZE dalle famiglie
    }

    // Estrai i corsi destinatari del turno
    let destinatari = listaID(dati[i][cDest]);

    // Verifica se l'utente è interessato a questo turno
    let trovato = false;

    if(vedeTutto){
      // Se vedeTutto=true (admin staff), mostra comunque il turno
      trovato = true;
    } else {
      // Verifica se almeno uno dei corsi dell'utente è nei destinatari.
      // La colonna può contenere l'ID oppure il nome del corso.
      trovato = destinatari.some(function(destinatario){
        return destinatarioCorrisponde(destinatario, corsi);
      });
    }

    if(!trovato) continue;  // Non rilevante per questo utente

    // Turno include, aggiungi ai risultati
    let coloreAttivita = String(coloriAttivita[i - 1][0] || "").toLowerCase();
    risultati.push({
      data: formattaData(dati[i][cData]),
      orario: dati[i][cOra],
      cosa: dati[i][cCosa],
      colore: coloreAttivita === "#ffffff" ? "" : coloreAttivita,
      destinatari: destinatari,
      stato: stato,  // BOZZA o PUBBLICATO
      note: cNote === -1 ? "" : dati[i][cNote],
      versione: cVersione === -1 ? 1 : Number(dati[i][cVersione] || 1)
    });
  }
  return risultati;
}

/**
 * Converte una data in formato leggibile italiano
 * @param {Date|string} data - Data da formattare
 * @returns {string} Data formattata come "GG/MM/AAAA"
 */
function formattaData(data){
  if(Object.prototype.toString.call(data) == "[object Date]"){
    return Utilities.formatDate(data, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return data;  // Se non è Date, ritorna così com'è
}

/****************************************************
 NOTIFICHE VERSIONI - Tracker delle letture
 Foglio VERSIONI: quale versione ogni utente ha visto
****************************************************/

/**
 * Recupera quale versione calendario l'utente ha già visto
 * Usato per mostrare alert "Nuovo aggiornamento" se versione > versioneVista
 * 
 * @param {string} email - Email dell'utente
 * @returns {number} Versione vista (0 se mai visto)
 */
function getVersioneVista(email){
  const sh = getFoglio("VERSIONI");
  if(!sh) return 0;  // Foglio VERSIONI non esiste

  const dati = sh.getDataRange().getValues();
  if(dati.length < 2) return 0;  // Foglio vuoto

  const h = dati[0];
  const cEmail = trovaColonna(h, "Email");
  const cVersioneVista = trovaColonna(h, "Versione Vista");

  // Colonne non trovate
  if (cEmail === -1 || cVersioneVista === -1) return 0;

  // Cerca email nella traccia
  for(let i=1; i<dati.length; i++){
    if(String(dati[i][cEmail]).trim().toLowerCase() == email.toLowerCase()){
      return Number(dati[i][cVersioneVista] || 0);
    }
  }
  return 0;  // Email non trovata = mai visto
}

/**
 * Registra che un utente ha visto una versione del calendario
 * Chiamato quando l'utente clicca il pulsante di conferma
 * 
 * @param {string} email - Email dell'utente
 * @param {number} versione - Versione che ha visto
 */
function salvaVersioneVista(email, versione){
  const sh = getFoglio("VERSIONI");
  if(!sh) return;  // Foglio non esiste, skip

  const dati = sh.getDataRange().getValues();
  const h = dati[0];
  const cEmail = trovaColonna(h, "Email");
  const cVersioneVista = trovaColonna(h, "Versione Vista");

  // Se colonne non trovate, crea la struttura
  if (cEmail === -1 || cVersioneVista === -1) {
    sh.clear();
    sh.appendRow(["Email", "Versione Vista"]);
    sh.appendRow([email, versione]);
    return;
  }

  // Aggiorna email esistente
  for(let i=1; i<dati.length; i++){
    if(String(dati[i][cEmail]).trim().toLowerCase() == email.toLowerCase()){
      sh.getRange(i+1, cVersioneVista+1).setValue(versione);
      return;
    }
  }

  // Email non trovata, aggiungi nuova riga
  let nuovaRiga = [];
  nuovaRiga[cEmail] = email;
  nuovaRiga[cVersioneVista] = versione;
  
  sh.appendRow(nuovaRiga);
}

/**
 * COMANDO AMMINISTRATORE: Pubblica tutti i turni BOZZA come PUBBLICATO
 * Questo trigger notifiche a tutte le famiglie
 * 
 * Processo:
 * 1. Incrementa versione globale calendario
 * 2. Aggiorna colonna "Versione" per tutti i turni con stato="PUBBLICATO"
 * 3. Salva timestamp dell'ultima pubblicazione
 * 4. Mostra alert di conferma
 * 
 * Nota: Docenti NON possono eseguire questo (limite di sicurezza)
 * Solo ADMIN o ruoli con permesso "Pubblica"
 */
function pubblicaAggiornamenti(){
  let versione = nuovaVersione();  // Incrementa versione globale
  const sh = getFoglio("CALENDARIO");
  const dati = sh.getDataRange().getValues();
  const h = dati[0];

  const cStato = trovaColonna(h, "Stato");
  const cVersione = trovaColonna(h, "Versione");

  // Aggiorna versione solo per turni PUBBLICATO
  // (Così le famiglie sapranno che c'è un aggiornamento)
  for(let i=1;i<dati.length;i++){
    let stato = String(dati[i][cStato]).trim().toUpperCase();
    if(stato == "PUBBLICATO"){
      sh.getRange(i+1, cVersione+1).setValue(versione);  // Marca il turno
    }
  }

  // Registra timestamp della pubblicazione
  PropertiesService
  .getScriptProperties()
  .setProperty("ULTIMA_PUBBLICAZIONE", new Date().toString());

  // Notifica l'admin
  SpreadsheetApp
  .getUi()
  .alert("📢 Aggiornamento pubblicato\nVersione " + versione);
}

/**
 * Salva un nuovo materiale didattico nel foglio MATERIALI
 * Chiamato dai docenti/admin tramite form upload nel portale
 * Il materiale viene creato con Pubblicato="SI" (visibile subito)
 * 
 * @param {string} idCorso - Codice corso
 * @param {string} titolo - Nome del file
 * @param {string} descrizione - Note per gli allievi
 * @param {string} linkDrive - URL Google Drive
 * @returns {Object} {errore, messaggio}
 */
function salvaNuovoMateriale(idCorso, titolo, descrizione, linkDrive) {
  const sh = getFoglio("MATERIALI");
  if(!sh) return { errore: true, messaggio: "Foglio MATERIALI non trovato." };
  
  const dati = sh.getDataRange().getValues();
  const h = dati[0];

  // Mappa colonne
  const cID = trovaColonna(h, "ID CORSI");
  const cTitolo = trovaColonna(h, "Titolo");
  const cDesc = trovaColonna(h, "Descrizione");
  const cLink = trovaColonna(h, "Link Drive");
  const cPubblicato = trovaColonna(h, "Pubblicato");
  const cData = trovaColonna(h, "Data");

  // Validazione: colonne obbligatorie
  if (cID === -1 || cTitolo === -1 || cLink === -1) {
    return { errore: true, messaggio: "Struttura delle colonne del foglio MATERIALI non valida." };
  }

  // Componi nuova riga (sparse array)
  let nuovaRiga = [];
  nuovaRiga[cID] = idCorso;
  nuovaRiga[cTitolo] = titolo;
  if (cDesc !== -1) nuovaRiga[cDesc] = descrizione;  // Opzionale
  nuovaRiga[cLink] = linkDrive;
  
  // Pubblica subito
  if (cPubblicato !== -1) nuovaRiga[cPubblicato] = "SI";
  
  // Registra data di caricamento
  if (cData !== -1) {
    nuovaRiga[cData] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  
  // Aggiungi al foglio
  sh.appendRow(nuovaRiga);
  return { errore: false, messaggio: "Materiale caricato e salvato correttamente!" };
}

/****************************************************
 CARICAMENTO GLOBALE PORTALE - Aggregazione dati
 Funzione principale chiamata dal frontend
****************************************************/

/**
 * Controlla se l'utente ha nuovi aggiornamenti da leggere
 * True se: versioneVista < versione attuale
 * Usato per mostrare alert alle famiglie
 * 
 * @param {string} email - Email dell'utente
 * @returns {boolean} true se ci sono aggiornamenti da leggere
 */
function controlloAggiornamenti(email){
  return (getVersioneVista(email) < getVersione());
}

/**
 * Recupera timestamp dell'ultima pubblicazione
 * Visualizzato nel portale per info trasparenza
 * 
 * @returns {string} Data/ora formattata (es: "29/08/2026 14:30") o vuoto se mai pubblicato
 */
function getUltimaPubblicazione(){
  let d = PropertiesService.getScriptProperties().getProperty("ULTIMA_PUBBLICAZIONE");
  if(!d) return "";
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
}

/**
 * FUNZIONE PRINCIPALE DEL PORTALE - Aggregazione completa dati utente
 * Chiamata dal frontend via google.script.run.caricaPortale(email)
 * 
 * Ritorna un oggetto "state" che contiene:
 * - Autenticazione e profilo utente
 * - Calendario filtrato per i corsi dell'utente
 * - Stato notifiche aggiornamenti
 * - Timestamp ultima pubblicazione
 * 
 * @param {string} email - Email dell'utente che accede
 * @returns {Object} 
 *   - Errore: {errore: true, messaggio: "..."}
 *   - Successo: {utente, corsiVisualizzati, prove, versione, versioneVista, aggiornamenti, ultimaPubblicazione}
 */
function caricaPortale(email){
  // STEP 1: Autenticazione - determina tipo utente (staff vs famiglia)
  let utente = getUtente(email);
  if(utente.errore) return utente;  // Email non riconosciuta

  let corsi = utente.corsi || [];
  
  // STEP 2: Carica calendario filtrato per questo utente
  let calendario = getCalendario(corsi, utente.vedeTutto || false, utente.tipo);

  // STEP 3: Assembla stato completo del portale
  return {
    utente: utente,  // Profilo utente con permessi
    corsiVisualizzati: creaCorsiVisualizzati(corsi),  // Nomi leggibili corsi
    orariLezioni: getOrariLezioni(corsi),
    prove: calendario,  // Turni filtrati
    versione: getVersione(),  // Versione attuale calendar
    versioneVista: getVersioneVista(email),  // Versione che ha visto l'utente
    aggiornamenti: utente.tipo !== "staff" ? controlloAggiornamenti(email) : false,  // Alert per famiglie e individuali
    ultimaPubblicazione: getUltimaPubblicazione()  // Timestamp info
  };
}

/**
 * Marca che l'utente ha visto gli aggiornamenti
 * Toglie l'alert rosso al prossimo login
 * Chiamata dal pulsante "✅ Ho letto e preso visione"
 * 
 * @param {string} email - Email dell'utente
 * @returns {boolean} true
 */
function confermaVisione(email){
  salvaVersioneVista(email, getVersione());
  return true;
} 