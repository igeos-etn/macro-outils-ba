// ==============================================================================
// 1. CONFIGURATION GÉNÉRALE
// ==============================================================================

const VERSION_SCRIPT = "0.9.18"; // Ajout de la rotation des logs
const AGENCE_ID = "AG98"; // <--- MODIFIEZ ICI LE NUMERO D'AGENCE
const AUTHORIZED_GROUP = "transfert-boite-agence-98@etn.fr";
const LOG_FILE_NAME = "LOGS – Script Boite Agence"; // Nom du fichier de logs
const LOG_SHEET_NAME = "Logs";
const EMAIL_ADMIN_ALERTE = "alertes@etn.fr";

const MAX_LOG_ROWS = 250000; // 250 000 lignes * 4 colonnes = 1 000 000 de cellules

// ==============================================================================
// 2. POINT D'ENTRÉE (WEB APP)
// ==============================================================================

function doGet() {
  const userEmail = Session.getActiveUser().getEmail();

  // 1. Vérification autorisation (avec Cache)
  if (!isUserAuthorized(userEmail)) {
    logWarn(`Acces refuse (PRE-LOCK) pour ${userEmail}`);
    return reponseJSON(userEmail, false, "Acces refuse. Non autorise.");
  }

  // 2. Verrouillage
  logInfo(`Acces autorise pour ${userEmail}`);
  const lock = LockService.getScriptLock();

  try {
    if (!lock.tryLock(30000)) {
      logWarn(`Tentative avortee (Verrouille) pour ${userEmail}`);
      return reponseJSON(userEmail, true, "Script deja en cours. Reesayez plus tard.");
    }

    logInfo(`Script lance (v${VERSION_SCRIPT} - ${AGENCE_ID}) par ${userEmail}`);

    // 3. Traitement
    const rapport = traitementMails(userEmail);

    // Construction réponse
    let response = {
      version: VERSION_SCRIPT,
      agence: AGENCE_ID,
      user: userEmail,
      authorized: true,
      message: "Traitement termine.",
      transfertStatus: rapport.transfertStatus,
      messageTransfert: rapport.nbre,
      nbrMsgTransfert: rapport.nbrMsgTransfert,
      msgTraites: rapport.msgTraites
    };

    return ContentService
      .createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    logError("Erreur critique script : " + e.message);
    return reponseJSON(userEmail, true, "Erreur interne : " + e.message);
  } finally {
    lock.releaseLock();
    logInfo("Fin execution - Verrou relache.");
  }
}

function reponseJSON(user, auth, msg) {
  return ContentService.createTextOutput(JSON.stringify({
    version: VERSION_SCRIPT, agence: AGENCE_ID, user: user, authorized: auth, message: msg
  })).setMimeType(ContentService.MimeType.JSON);
}

// ==============================================================================
// 3. TRAITEMENT DES EMAILS
// ==============================================================================

function traitementMails(userEmail) {
  const ADRESSE_DESTINATAIRE = userEmail;
  const libelleUtilisateur = verifLibelle(userEmail);
  const LIMITE_POIDS_GMAIL = 24 * 1024 * 1024; // 24 Mo

  let compteurTraites = 0;
  let msgTraites = [];

  const compteScript = Session.getEffectiveUser().getEmail();
  const threads = GmailApp.getInboxThreads(0, 20);

  if (threads.length > 0) {
    logInfo(`Debut du traitement de ${threads.length} threads...`);

    for (const thread of threads) {
      const messages = thread.getMessages();
      let messagesTraitesDansCeThread = false;

      for (const message of messages) {
        if (!message.isUnread()) continue;

        const messageId = message.getId();

        if (message.getFrom().indexOf(compteScript) !== -1) {
          logInfo(`[Ignoré] Message sortant détecté.`);
          message.markRead();
          continue;
        }

        const sujetOriginal = message.getSubject();
        const from = message.getFrom();
        const expediteurOriginal = message.getFrom().replace(/[<]/g, "&lt;").replace(/[>]/g, "&gt;");
        const dateOriginale = message.getDate();
        const to = message.getTo().replace(/[<]/g, "&lt;").replace(/[>]/g, "&gt;");
        const cc = message.getCc().replace(/[<]/g, "&lt;").replace(/[>]/g, "&gt;");
        const ligneCc = (cc && cc.length > 0) ? `<strong>Cc : </strong> ${cc}<br>` : "";

        let piecesJointesAEnvoyer = [];
        let mapInlineImages = {};
        let poidsTotal = 0;

        const triAttachments = getInlineImagesFromRaw(message);
        piecesJointesAEnvoyer = triAttachments.regular;
        mapInlineImages = triAttachments.inline;

        piecesJointesAEnvoyer.forEach(pj => {
          try { poidsTotal += (typeof pj.getSize === 'function') ? pj.getSize() : pj.getBytes().length; } catch(e) {}
        });
        for (let cid in mapInlineImages) {
          let img = mapInlineImages[cid];
          try { poidsTotal += (typeof img.getSize === 'function') ? img.getSize() : img.getBytes().length; } catch(e) {}
        }

        let htmlSupplementaire = "";

        if (poidsTotal > LIMITE_POIDS_GMAIL) {
          logInfo(`[ID:${messageId}] PJ lourdes (${(poidsTotal / 1e6).toFixed(1)}Mo). Mode Drive.`);
          try {
            let tout = piecesJointesAEnvoyer.concat(Object.values(mapInlineImages));
            htmlSupplementaire = uploadToDriveAndGetLinks(tout, sujetOriginal);
            piecesJointesAEnvoyer = [];
            mapInlineImages = {};
          } catch (errDrive) {
            traiterErreurBloquante(msgTraites, userEmail, compteScript, sujetOriginal, messageId, "Echec Drive PJ: " + errDrive.message);
            continue;
          }
        }

        let enteteTransfert = `
          <div style="border-left: 3px solid #ccc; padding-left: 10px; margin-bottom: 15px; font-family: Arial, sans-serif;">
            <strong>---------- Message transfere ----------</strong><br>
            <strong>De :</strong> ${expediteurOriginal}<br>
            <strong>Date :</strong> ${dateOriginale.toLocaleString()}<br>
            <strong>À : </strong> ${to}<br>
            <strong>Objet :</strong> ${sujetOriginal}<br>
            ${ligneCc}
          </div>`;

        const corpsOriginal = message.getBody();
        const estProbablementHtml = /<[a-z][\s\S]*>/i.test(corpsOriginal);
        const corpsFinal = estProbablementHtml ? corpsOriginal : `<div style="white-space: pre-wrap; font-family: Arial, sans-serif;">${corpsOriginal}</div>`;

        const nouveauCorpsHtml = enteteTransfert + htmlSupplementaire + corpsFinal;

        let sujetClean = sujetOriginal;
        if (sujetClean.length > 240) { sujetClean = sujetClean.substring(0, 240) + "..."; }
        const nouveauSujet = `Fwd: ${sujetClean}`;

        try {
          GmailApp.sendEmail(ADRESSE_DESTINATAIRE, nouveauSujet, "Necessite client HTML.", {
            htmlBody: nouveauCorpsHtml,
            attachments: piecesJointesAEnvoyer,
            inlineImages: mapInlineImages,
            "replyTo": from
          });

          compteurTraites++;
          logSuccess(msgTraites, sujetOriginal, messageId, htmlSupplementaire);
          message.markRead();
          messagesTraitesDansCeThread = true;

        } catch (e) {
          const analyse = execeptionMsg(e);

          if (analyse.isBodySizeError) {
            logInfo(`[ID:${messageId}] Corps lourd/Format incorrect. Mode EML.`);
            try {
              const blobEml = Utilities.newBlob(message.getRawContent(), 'message/rfc822', 'message_complet.eml');
              const toutSurDrive = piecesJointesAEnvoyer.concat([blobEml]).concat(Object.values(mapInlineImages));
              const liensDrive = uploadToDriveAndGetLinks(toutSurDrive, sujetOriginal);
              const explicationFr = "Ce mail est mal formaté ou trop volumineux, il est donc envoyé sous forme de lien Drive.";
              const explicationEn = "This email is malformed or too large, so it is sent as a Drive link.";

              const corpsAlerteUser = `
                <div style="font-family: Arial, sans-serif; border: 1px solid #ccc; padding: 15px; background-color: #fff3cd;">
                  ${enteteTransfert}
                  <br>
                <strong>⚠️ INFORMATION (FR) :</strong> ${explicationFr}<br>
                <strong>⚠️ INFORMATION (EN) :</strong> ${explicationEn}<br>
                <br>
                  ${liensDrive}
                </div>`;

              GmailApp.sendEmail(ADRESSE_DESTINATAIRE, nouveauSujet + " [Format .eml]", "Voir liens Drive.", {
                htmlBody: corpsAlerteUser, "replyTo": from
              });

              compteurTraites++;
              logSuccess(msgTraites, sujetOriginal, messageId, "Via Drive (Mode EML)");
              message.markRead();
              messagesTraitesDansCeThread = true;

            } catch (errSauvetage) {
              const msgErr = "Echec Mode EML/Drive : " + errSauvetage.message;
              traiterErreurBloquante(msgTraites, userEmail, compteScript, sujetOriginal, messageId, msgErr);
            }
          } else {
            const msgErr = analyse.messageClair;
            msgTraites.push({ id: messageId, sujet: sujetOriginal, status: "bloque", erreur: "true", msgErreur: msgErr });

            if (!analyse.stopScript) {
              envoyerNotificationEchec(userEmail, compteScript, sujetOriginal, messageId, msgErr);
            }
            if (analyse.stopScript) break;
          }
        }
      }

      if (messagesTraitesDansCeThread) thread.addLabel(libelleUtilisateur);
      if (thread.isUnread() === false) thread.moveToArchive();
    }

    return { transfertStatus: (compteurTraites > 0), nbre: compteurTraites, nbrMsgTransfert: compteurTraites + "/" + threads.length, msgTraites: msgTraites };
  } else {
    logInfo("Aucun e-mail a traiter.");
    return { transfertStatus: false, nbre: 0, nbrMsgTransfert: "0/0", msgTraites: [] };
  }
}

function traiterErreurBloquante(rapport, user, compteScript, sujet, id, msgErr) {
  logError(`[ID:${id}] ${msgErr}`);
  rapport.push({ id: id, sujet: sujet, status: "bloque", erreur: "true", msgErreur: msgErr });
  envoyerNotificationEchec(user, compteScript, sujet, id, msgErr);
}

// ==============================================================================
// 4. DRIVE, FICHIERS ET EXTRACTION (HACK NO-API)
// ==============================================================================

function getInlineImagesFromRaw(message) {
  const attachments = message.getAttachments();
  const rawContent = message.getRawContent();
  const htmlBody = message.getBody();
  const inlineMap = {};
  const regularAttachments = [];
  const fileNameToCid = {};

  if (!attachments || attachments.length === 0) return { inline: {}, regular: [] };

  const referencedCids = {};
  const cidRegex = /src=["']cid:([^"']+)["']/gi;
  let m;
  while ((m = cidRegex.exec(htmlBody)) !== null) { referencedCids[m[1]] = true; }

  const parts = rawContent.split(/\r?\n--/);
  parts.forEach(part => {
    const headerBlock = part.split(/\r?\n\r?\n/)[0];
    const idMatch = headerBlock.match(/Content-ID:\s*<([^>]+)>/i) || 
                    headerBlock.match(/Content-ID:\s*([^; \s\r\n]+)/i) ||
                    headerBlock.match(/X-Attachment-Id:\s*([^; \s\r\n]+)/i);
    const nameMatch = headerBlock.match(/filename="?([^";\n\r]+)"?/i) || 
                      headerBlock.match(/name="?([^";\n\r]+)"?/i);
    
    if (idMatch && nameMatch) {
      let cid = idMatch[1];
      let fileName = decodeMimeString(nameMatch[1].trim());
      fileNameToCid[fileName] = cid;
    }
  });

  const imgRegex = /<img[^>]+(?:src=["']cid:([^"']+)["'][^>]+alt=["']([^"']+)["']|alt=["']([^"']+)["'][^>]+src=["']cid:([^"']+)["'])/gi;
  while ((m = imgRegex.exec(htmlBody)) !== null) {
    const cid = m[1] || m[4];
    let alt = m[2] || m[3];
    if (cid && alt) {
      alt = alt.replace(/&#(\d+);/g, (match, dec) => String.fromCharCode(dec))
               .replace(/&quot;/g, '"').replace(/&amp;/g, '&')
               .replace(/&lt;/g, '<').replace(/&gt;/g, '>');
      if (!fileNameToCid[alt]) { fileNameToCid[alt] = cid; }
    }
  }

  attachments.forEach(pj => {
    const pjName = pj.getName();
    const cid = fileNameToCid[pjName];
    if (cid && referencedCids[cid]) {
      inlineMap[cid] = pj;
    } else {
      const cleanBlob = Utilities.newBlob(pj.getBytes(), pj.getContentType(), pjName);
      regularAttachments.push(cleanBlob);
    }
  });

  return { inline: inlineMap, regular: regularAttachments };
}

function decodeMimeString(str) {
  if (!str || str.indexOf('=?') === -1) return str;
  try {
    return str.replace(/=\?([^?]+)\?([BQ])\?([^?]+)\?=/gi, (match, charset, encoding, text) => {
      if (encoding.toUpperCase() === 'B') {
        const decoded = Utilities.base64Decode(text);
        return Utilities.newBlob(decoded).getDataAsString(charset);
      } else if (encoding.toUpperCase() === 'Q') {
        let temp = text.replace(/_/g, ' ');
        temp = temp.replace(/=([0-9A-F]{2})/gi, (m, hex) => String.fromCharCode(parseInt(hex, 16)));
        return temp;
      }
      return match;
    });
  } catch (e) { return str; }
}

function uploadToDriveAndGetLinks(piecesJointes, sujetMessage) {
  const NOM_DOSSIER_RACINE = "BA-Pieces Jointes Transferees";
  const locale = Session.getActiveUserLocale();
  const estFrancophone = locale && locale.toLowerCase().startsWith('fr');

  const titre = estFrancophone ? "Pieces jointes (Dossier : {nom}) :" : "Attachments (Folder: {nom}):";
  const explication = estFrancophone ? "Fichiers securises sur le Google Drive de l'agence." : "Files saved to the agency's Google Drive.";

  let dossierRacine;
  const racines = DriveApp.getFoldersByName(NOM_DOSSIER_RACINE);
  dossierRacine = racines.hasNext() ? racines.next() : DriveApp.createFolder(NOM_DOSSIER_RACINE);

  const randomId = Math.random().toString(36).substring(2, 10).toUpperCase();
  let safeSujet = sujetMessage.substring(0, 40).replace(/[/\\?%*:|"<>\.]/g, '-');
  const nomSousDossier = `${safeSujet}... [${randomId}]`;
  const sousDossier = dossierRacine.createFolder(nomSousDossier);

  let htmlLinks = `
    <div style="background-color: #f8f9fa; padding: 12px; border: 1px solid #dadce0; border-radius: 6px; font-family: Arial;">
       <strong>${titre.replace("{nom}", nomSousDossier)}</strong><br>
       <em style="color: #5f6368; font-size: 0.9em;">${explication}</em><ul style="margin-top:10px;">`;

  for (const pj of piecesJointes) {
    const fichier = sousDossier.createFile(pj);
    fichier.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const tailleMo = (fichier.getSize() / 1024 / 1024).toFixed(2);
    htmlLinks += `<li style="margin-bottom: 5px;"><a href="${fichier.getUrl()}" target="_blank" style="color: #1a73e8; font-weight: bold;">${pj.getName()}</a> <span style="color:#70757a">(${tailleMo} Mo)</span></li>`;
  }
  return htmlLinks + `</ul></div><br>`;
}

// ==============================================================================
// 5. GESTION DES LOGS (OPTIMISÉ + ROTATION)
// ==============================================================================

let cachedLogSheet = null;

function getLogSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('LOG_SS_ID');
  
  if (ssId) {
    try {
      return SpreadsheetApp.openById(ssId);
    } catch(e) {
      props.deleteProperty('LOG_SS_ID');
    }
  }
  
  const files = DriveApp.getFilesByName(LOG_FILE_NAME);
  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(LOG_FILE_NAME);
  }
  
  props.setProperty('LOG_SS_ID', ss.getId());
  return ss;
}

function getGlobalLogSheet() {
  if (cachedLogSheet) return cachedLogSheet;
  
  const ss = getLogSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(["Timestamp", "Utilisateur", "Niveau", "Message"]);
  }
  
  cachedLogSheet = sheet;
  return sheet;
}

function archiverLogsSiBesoin(sheet) {
  // Vérifie si la limite des 250 000 lignes (1 million de cellules) est atteinte
  if (sheet.getLastRow() >= MAX_LOG_ROWS) {
    const ss = sheet.getParent(); // Récupère le tableur parent
    
    // On ajoute la date et l'heure pour garantir un nom de fichier unique
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
    ss.rename(`${LOG_FILE_NAME} - ARCHIVE ${dateStr}`);
    
    // On force l'oubli de ce tableur pour que le script en crée un tout neuf
    PropertiesService.getScriptProperties().deleteProperty('LOG_SS_ID');
    cachedLogSheet = null; 
    
    // On génère et on récupère la nouvelle feuille de logs
    const nouvelleSheet = getGlobalLogSheet();
    nouvelleSheet.appendRow([new Date(), "SYSTEM", "INFO", `Nouveau fichier de log créé (Ancien archivé le ${dateStr})`]);
    return nouvelleSheet;
  }
  return sheet;
}

function writeLog(level, msg) {
  const row = [new Date(), Session.getActiveUser().getEmail(), level, msg];
  try { 
    // On charge la feuille, on vérifie si elle doit être archivée, puis on écrit
    let sheet = getGlobalLogSheet();
    sheet = archiverLogsSiBesoin(sheet);
    sheet.appendRow(row); 
  } catch (e) { 
    Logger.log("Err Log Global: " + e); 
  }
}

function logInfo(msg) { writeLog("INFO", msg); }
function logWarn(msg) { writeLog("WARN", msg); }
function logError(msg) { writeLog("ERROR", msg); }

// ==============================================================================
// 6. UTILITAIRES DIVERS
// ==============================================================================

function execeptionMsg(e) {
  let errStr = "";
  try { errStr = LanguageApp.translate(e.toString(), '', 'en').toLowerCase(); } catch (err) { errStr = e.toString().toLowerCase(); }

  let res = { messageClair: "Erreur technique : " + e.message, stopScript: false, isBodySizeError: false };

  if (errStr.includes("attachment") && errStr.includes("size")) {
    res.messageClair = "Non transfere : Pieces jointes > 25Mo";
  } else if ((errStr.includes("body") && errStr.includes("size")) || (errStr.includes("taille") && errStr.includes("corps"))) {
    res.messageClair = "Non transfere : Contenu mail trop volumineux";
    res.isBodySizeError = true;
  } else if (errStr.includes("limit") && errStr.includes("exceeded")) {
    res.messageClair = "STOP URGENT : Quota Google depasse";
    res.stopScript = true;
  }
  return res;
}

function isUserAuthorized(email) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "auth_" + email;
  const cachedResult = cache.get(cacheKey);

  if (cachedResult !== null) {
    return cachedResult === "true";
  }

  try { 
    const isAuth = GroupsApp.getGroupByEmail(AUTHORIZED_GROUP).hasUser(email);
    
    // Modification ici : durée du cache variable selon le résultat
    if (isAuth) {
      cache.put(cacheKey, "true", 7200); // 2 heures si autorisé
    } else {
      cache.put(cacheKey, "false", 60);   // 1 minute seulement si refusé
    }
    
    return isAuth;
  } catch (e) { 
    logError("Err Groupe: " + e.message); 
    return false; 
  }
}

function verifLibelle(userEmail) {
  const label = GmailApp.getUserLabelByName(userEmail);
  return label ? label : GmailApp.createLabel(userEmail);
}

function envoyerNotificationEchec(userExecutant, compteScript, sujet, id, err) {
  const sujetAlerte = `[alertes] Erreur Script BA - ${AGENCE_ID}`;
  const corps =
    "Bonjour l'equipe technique,\n\n" +
    "Une erreur bloquante a empeche le transfert d'un email.\n\n" +
    "--- CONTEXTE ---\n" +
    "Agence         : " + AGENCE_ID + "\n" +
    "Version      : " + VERSION_SCRIPT + "\n" +
    "Utilisateur  : " + userExecutant + "\n" +
    "Compte Script: " + compteScript + "\n\n" +
    "--- DETAILS DU MESSAGE ---\n" +
    "ID Message   : " + id + "\n" +
    "Sujet        : " + sujet + "\n\n" +
    "--- ERREUR TECHNIQUE ---\n" +
    "Raison       : " + err + "\n\n" +
    "Le message est reste dans la boite de reception.";

  try { GmailApp.sendEmail(EMAIL_ADMIN_ALERTE, sujetAlerte, corps); } catch (e) { }
}

function logSuccess(tab, sujet, id, extra) {
  tab.push({ id: id, sujet: sujet, status: "transfere", erreur: "false", msgErreur: extra ? "Via Drive" : "" });
  logInfo(`OK : "${sujet}" (ID:${id}) transfere.`);
}
