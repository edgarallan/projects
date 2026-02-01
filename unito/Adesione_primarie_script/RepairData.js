/**
 * @fileoverview Script di utilità per riparare i dati su Firebase
 * sincronizzando le assegnazioni "SI/NO" che sono sul foglio ma non nel nodo permanent.
 */

var RepairDataService = (function() {

  /**
   * Forza la sincronizzazione di tutte le "Proposte accettate" presenti nel foglio
   * verso il nodo permanent di Firebase, assicurandosi di generare ID se mancano.
   */
  function repairAssignmentsFromSheet() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('RIPARAZIONE DATI (PRIMARIE)', 
      'Questo script leggerà TUTTE le righe del foglio PRIMARIE. ' +
      'Se una riga ha "Proposta accettata" (SI/NO) ma non è su Firebase, verrà aggiunta. ' +
      'Continuare?', ui.ButtonSet.YES_NO);
    
    if (response !== ui.Button.YES) return;

    try {
      const ssPri = Utilities_.findSpreadsheetByName(CONFIG.REPOPULATE_SOURCE_FILE_PRIMARY);
      const sheet = ssPri.getSheetByName(CONFIG.REPOPULATE_SOURCE_SHEET_NAME);
      const data = sheet.getDataRange().getValues();
      const headers = data.shift();

      const idColIdx = headers.indexOf('firebase_id');
      const acceptedColIdx = headers.indexOf('Proposta accettata');
      const tsColIdx = headers.indexOf('Informazioni cronologiche');
      const emailColIdx = headers.indexOf('Indirizzo email');

      Logger.log('Caricamento record esistenti da Firebase...');
      const fbPrim = FirebaseService.firebaseGet(CONFIG.REQUESTS_NODE_PRIMARY) || {};
      const fbSec = FirebaseService.firebaseGet(CONFIG.REQUESTS_NODE_SECONDARY) || {};
      const firebaseMap = new Map();
      
      const populateMap = (data) => {
        for (const id in data) {
          const r = data[id];
          if (r.timestamp && r.email) {
            const ts = new Date(r.timestamp).toLocaleString('it-IT', { timeZone: CONFIG.SCRIPT_TIMEZONE });
            firebaseMap.set(`${ts}_${r.email.trim()}`, id);
          }
        }
      };
      populateMap(fbPrim);
      populateMap(fbSec);

      let newIdsGenerated = 0;
      data.forEach((row, index) => {
        let fid = String(row[idColIdx] || '').trim();
        const email = String(row[emailColIdx] || '').trim();
        const timestampRaw = row[tsColIdx];
        const accepted = String(row[acceptedColIdx] || '').trim().toUpperCase();
        
        if (!email || !accepted || accepted === '') return;

        // Se manca il firebase_id nel foglio, ne generiamo uno
        if (!fid) {
          const tsSheet = new Date(timestampRaw).toLocaleString('it-IT', { timeZone: CONFIG.SCRIPT_TIMEZONE });
          const key = `${tsSheet}_${email}`;
          const existingId = firebaseMap.get(key);

          if (existingId) {
            fid = existingId;
            Logger.log(`Riga ${index + 2}: Recuperato ID esistente ${fid} per ${email}`);
          } else {
            fid = Utilities_.generateUniqueId();
            Logger.log(`Riga ${index + 2}: Generazione nuovo ID ${fid} per ${email}`);
            newIdsGenerated++;
          }
          
          const cell = sheet.getRange(index + 2, idColIdx + 1);
          cell.setDataValidation(null);
          cell.setValue(fid);
        }
      });

      // Eseguiamo il sync standard
      const stats = SheetSyncService.syncProposalsFromFile(ssPri, CONFIG.ASSEGNAZIONI_PRIMARIE_NODE);

      ui.alert(`Riparazione completata!\n\nNuovi ID generati: ${newIdsGenerated}\nAssegnazioni syncate: ${stats.successCount}`);
      
    } catch (e) {
      Logger.log('ERRORE in repairAssignmentsFromSheet: ' + e.toString());
      ui.alert('Errore: ' + e.message);
    }
  }

  return {
    repairAssignmentsFromSheet: repairAssignmentsFromSheet
  };

})();
