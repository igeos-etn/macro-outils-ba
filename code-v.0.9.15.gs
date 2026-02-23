// ==============================================================================
// 1. CONFIGURATION GÉNÉRALE
// ==============================================================================

const VERSION_SCRIPT = "0.9.15";
const AGENCE_ID = "AG98"; // <--- MODIFIEZ ICI LE NUMERO D'AGENCE
const AUTHORIZED_GROUP = "transfert-boite-agence-98@etn.fr";
const LOG_FILE_NAME = "LOGS – Script Boite Agence"; // Nom du fichier de logs
const LOG_SHEET_NAME = "Logs";
const EMAIL_ADMIN_ALERTE = "alertes@etn.fr";

// ==============================================================================
// 2. POINT D'ENTRÉE (WEB APP)
// ==============================================================================

function doGet() {
  const userEmail = Session.getActiveUser().getEmail(); // L'utilisateur qui clique

  // 1. Vérification autorisation
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

// Helper pour réponse JSON rapide
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

    // --- BOUCLE SUR LES FILS DE DISCUSSION ---
    for (const thread of threads) {

      const messages = thread.getMessages();
      let messagesTraitesDansCeThread = false;

      // --- BOUCLE SUR LES MESSAGES DU FIL ---
      for (const message of messages) {

        // 1. Si déjà lu, on passe
        if (!message.isUnread()) {
          continue;
        }

        const messageId = message.getId();

        // 2. Anti-boucle infinie
        if (message.getFrom().indexOf(compteScript) !== -1) {
          logInfo(`[Ignoré] Message sortant détecté.`);
          message.markRead();
          continue;
        }

        // --- DÉBUT TRAITEMENT ---
        const sujetOriginal = message.getSubject();
        const from = message.getFrom();
        const expediteurOriginal = message.getFrom().replace(/[<]/g, "&lt;").replace(/[>]/g, "&gt;");
        const dateOriginale = message.getDate();
        const to = message.getTo().replace(/[<]/g, "&lt;").replace(/[>]/g, "&gt;");
        const cc = message.getCc().replace(/[<]/g, "&lt;").replace(/[>]/g, "&gt;");
        const ligneCc = (cc && cc.length > 0) ? `<strong>Cc : </strong> ${cc}<br>` : "";

        // --- GESTION PJ & INLINE (SANS API) ---
        let piecesJointesAEnvoyer = [];
        let mapInlineImages = {};
        let poidsTotal = 0;

        // Analyse du Raw Content pour trouver les CIDs
        const triAttachments = getInlineImagesFromRaw(message);

        piecesJointesAEnvoyer = triAttachments.regular;
        mapInlineImages = triAttachments.inline;

        // Calcul du poids total (Fichiers + Images Inline)
        piecesJointesAEnvoyer.forEach(pj => {
          try {
            poidsTotal += (typeof pj.getSize === 'function') ? pj.getSize() : pj.getBytes().length;
          } catch(e) { /* Fallback */ }
        });
        for (let cid in mapInlineImages) {
          let img = mapInlineImages[cid];
          try {
            poidsTotal += (typeof img.getSize === 'function') ? img.getSize() : img.getBytes().length;
          } catch(e) { /* Fallback */ }
        }

        let htmlSupplementaire = "";

        // --- CAS 1 : PJ TROP LOURDES ---
        if (poidsTotal > LIMITE_POIDS_GMAIL) {
          logInfo(`[ID:${messageId}] PJ lourdes (${(poidsTotal / 1e6).toFixed(1)}Mo). Mode Drive.`);
          try {
            // On combine tout pour le Drive
            let tout = piecesJointesAEnvoyer.concat(Object.values(mapInlineImages));
            htmlSupplementaire = uploadToDriveAndGetLinks(tout, sujetOriginal);

            // On vide les tableaux pour ne pas les envoyer par mail
            piecesJointesAEnvoyer = [];
            mapInlineImages = {};
          } catch (errDrive) {
            traiterErreurBloquante(msgTraites, userEmail, compteScript, sujetOriginal, messageId, "Echec Drive PJ: " + errDrive.message);
            continue;
          }
        }

        // Construction En-tête HTML
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
        // Si le corps ne contient pas de balises HTML probables, on préserve le formatage texte
        const estProbablementHtml = /<[a-z][\s\S]*>/i.test(corpsOriginal);
        const corpsFinal = estProbablementHtml ? corpsOriginal : `<div style="white-space: pre-wrap; font-family: Arial, sans-serif;">${corpsOriginal}</div>`;

        const nouveauCorpsHtml = enteteTransfert + htmlSupplementaire + corpsFinal;

        // Tronquer le sujet si trop long (> 240 chars) pour éviter erreurs Gmail
        let sujetClean = sujetOriginal;
        if (sujetClean.length > 240) { sujetClean = sujetClean.substring(0, 240) + "..."; }
        const nouveauSujet = `Fwd: ${sujetClean}`;

        // --- TENTATIVE D'ENVOI ---
        try {
          GmailApp.sendEmail(ADRESSE_DESTINATAIRE, nouveauSujet, "Necessite client HTML.", {
            htmlBody: nouveauCorpsHtml,
            attachments: piecesJointesAEnvoyer,
            inlineImages: mapInlineImages, // <--- C'est ici qu'on répare les images
            "replyTo": from
          });

          compteurTraites++;
          logSuccess(msgTraites, sujetOriginal, messageId, htmlSupplementaire);

          message.markRead();
          messagesTraitesDansCeThread = true;

        } catch (e) {
          // --- GESTION ERREUR ---
          const analyse = execeptionMsg(e);

          if (analyse.isBodySizeError) {
            // Mode EML de secours
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
            // Autres erreurs bloquantes
            const msgErr = analyse.messageClair;
            msgTraites.push({ id: messageId, sujet: sujetOriginal, status: "bloque", erreur: "true", msgErreur: msgErr });

            if (!analyse.stopScript) {
              envoyerNotificationEchec(userEmail, compteScript, sujetOriginal, messageId, msgErr);
            }
            if (analyse.stopScript) break;
          }
        }
      } // Fin boucle messages

      // Gestion archivage du Thread
      if (messagesTraitesDansCeThread) {
        thread.addLabel(libelleUtilisateur);
      }
      // On archive si plus rien de non lu
      if (thread.isUnread() === false) {
        thread.moveToArchive();
      }

    } // Fin boucle threads

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

/**
 * Extrait les images inline en analysant la structure MIME et le corps HTML.
 * Gère les cas où le Content-ID est absent ou remplacé par X-Attachment-Id,
 * et utilise l'attribut 'alt' du HTML comme dernier recours.
 */
function getInlineImagesFromRaw(message) {
  const attachments = message.getAttachments();
  const rawContent = message.getRawContent();
  const htmlBody = message.getBody();
  const inlineMap = {};
  const regularAttachments = [];
  const fileNameToCid = {};

  if (!attachments || attachments.length === 0) {
    return { inline: {}, regular: [] };
  }

  // 1. Identifier tous les CIDs réellement utilisés dans le corps HTML
  const referencedCids = {};
  const cidRegex = /src=["']cid:([^"']+)["']/gi;
  let m;
  while ((m = cidRegex.exec(htmlBody)) !== null) {
    referencedCids[m[1]] = true;
  }

  // 2. Analyse des parties MIME pour mapper les noms de fichiers aux CIDs
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

  // 3. Fallback : Mapper via l'attribut 'alt' du HTML si nécessaire
  const imgRegex = /<img[^>]+(?:src=["']cid:([^"']+)["'][^>]+alt=["']([^"']+)["']|alt=["']([^"']+)["'][^>]+src=["']cid:([^"']+)["'])/gi;
  while ((m = imgRegex.exec(htmlBody)) !== null) {
    const cid = m[1] || m[4];
    let alt = m[2] || m[3];
    if (cid && alt) {
      alt = alt.replace(/&#(\d+);/g, (match, dec) => String.fromCharCode(dec))
               .replace(/&quot;/g, '"').replace(/&amp;/g, '&')
               .replace(/&lt;/g, '<').replace(/&gt;/g, '>');
      if (!fileNameToCid[alt]) {
        fileNameToCid[alt] = cid;
      }
    }
  }

  // 4. Répartition finale : Uniquement en inline si le CID est référencé dans le HTML
  attachments.forEach(pj => {
    const pjName = pj.getName();
    const cid = fileNameToCid[pjName];

    if (cid && referencedCids[cid]) {
      inlineMap[cid] = pj;
    } else {
      // On crée un nouveau Blob pour "nettoyer" les métadonnées MIME d'origine
      // Cela force Gmail à traiter le fichier comme un nouvel attachment standard
      const cleanBlob = Utilities.newBlob(pj.getBytes(), pj.getContentType(), pjName);
      regularAttachments.push(cleanBlob);
    }
  });

  return { inline: inlineMap, regular: regularAttachments };
}

/**
 * Décode les chaînes de caractères encodées selon la RFC 2047 (ex: =?UTF-8?B?...)
 * Très courant dans les headers MIME pour les caractères accentués.
 */
function decodeMimeString(str) {
  if (!str || str.indexOf('=?') === -1) return str;

  try {
    return str.replace(/=\?([^?]+)\?([BQ])\?([^?]+)\?=/gi, (match, charset, encoding, text) => {
      if (encoding.toUpperCase() === 'B') {
        const decoded = Utilities.base64Decode(text);
        return Utilities.newBlob(decoded).getDataAsString(charset);
      } else if (encoding.toUpperCase() === 'Q') {
        // Quoted-Printable simplifié
        let temp = text.replace(/_/g, ' ');
        temp = temp.replace(/=([0-9A-F]{2})/gi, (m, hex) => String.fromCharCode(parseInt(hex, 16)));
        return temp;
      }
      return match;
    });
  } catch (e) {
    return str; // En cas d'échec, on renvoie la chaîne brute
  }
}

function uploadToDriveAndGetLinks(piecesJointes, sujetMessage) {
  const NOM_DOSSIER_RACINE = "BA-Pieces Jointes Transferees";
  const locale = Session.getActiveUserLocale();
  const estFrancophone = locale && locale.toLowerCase().startsWith('fr');

  const titre = estFrancophone ? "Pieces jointes (Dossier : {nom}) :" : "Attachments (Folder: {nom}):";
  const explication = estFrancophone
    ? "Fichiers securises sur le Google Drive de l'agence."
    : "Files saved to the agency's Google Drive.";

  let dossierRacine;
  const racines = DriveApp.getFoldersByName(NOM_DOSSIER_RACINE);
  dossierRacine = racines.hasNext() ? racines.next() : DriveApp.createFolder(NOM_DOSSIER_RACINE);

  const randomId = Math.random().toString(36).substring(2, 10).toUpperCase();
  // Sécurité nom dossier
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

    // getSize() marche toujours sur un objet File Drive
    const tailleMo = (fichier.getSize() / 1024 / 1024).toFixed(2);

    htmlLinks += `<li style="margin-bottom: 5px;"><a href="${fichier.getUrl()}" target="_blank" style="color: #1a73e8; font-weight: bold;">${pj.getName()}</a> <span style="color:#70757a">(${tailleMo} Mo)</span></li>`;
  }
  return htmlLinks + `</ul></div><br>`;
}

// ==============================================================================
// 5. GESTION DES LOGS
// ==============================================================================

function getLogSpreadsheet() {
  const files = DriveApp.getFilesByName(LOG_FILE_NAME);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  else return SpreadsheetApp.create(LOG_FILE_NAME);
}

function getGlobalLogSheet() {
  const ss = getLogSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(["Timestamp", "Utilisateur", "Niveau", "Message"]);
  }
  return sheet;
}

function getUserSpecificLogSheet() {
  const ss = getLogSpreadsheet();
  const userEmail = Session.getActiveUser().getEmail();
  const localPart = userEmail.split('@')[0].replace(/[^a-zA-Z0-9-._]/g, '_');
  const sheetName = LOG_SHEET_NAME + "_" + localPart;

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["Timestamp", "Utilisateur", "Niveau", "Message"]);
  }
  return sheet;
}

function writeLog(level, msg) {
  const row = [new Date(), Session.getActiveUser().getEmail(), level, msg];
  try { getGlobalLogSheet().appendRow(row); } catch (e) { Logger.log("Err Log Global: " + e); }
  try { getUserSpecificLogSheet().appendRow(row); } catch (e) { Logger.log("Err Log User: " + e); }
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
  try { return GroupsApp.getGroupByEmail(AUTHORIZED_GROUP).hasUser(email); }
  catch (e) { logError("Err Groupe: " + e.message); return false; }
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
