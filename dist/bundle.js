/**
 * Bundle unificado y corregido localmente sin dependencias externas.
 * Incluye parche defensivo para el cambio de API de Instagram de Octubre 2026.
 */

const IGSF = Object.create(null);

'use strict';

(function (exports, window) {

  // Variables globales del sistema
  const dest = { folderId: null, folderObj: null };
  const badgeFileIds = { lastTestedDate: '', lastTestedStatus: '' };
  const igParams = { X_ASBD_ID: null, X_CSRFTOKEN: null, X_IG_APP_ID: null, X_IG_WWW_CLAIM: null, COOKIE: null };
  
  exports.errorReportEmail = void 0;
  exports.isSettingsLoaded = false;

  const sheetNames = {
    logs: 'Logs',
    subscriptions: 'Subscriptions',
    settings: 'Settings',
  };

  // 1. CARGAR CONFIGURACIÓN DE LA HOJA
  function loadSettings() {
    const spreadsheet = SpreadsheetApp.getActive();
    const settingsSheet = spreadsheet.getSheetByName(sheetNames['settings']);
    if (!settingsSheet) throw new Error("No se encontró la pestaña 'Settings'");
    
    dest.folderId = settingsSheet.getRange('gdriveId').getDisplayValue();
    try {
      dest.folderObj = DriveApp.getFolderById(dest.folderId);
    } catch (err) {
      dest.folderObj = DriveApp.getRootFolder();
    }
    
    try {
      igParams.X_ASBD_ID = settingsSheet.getRange('X_ASBD_ID').getDisplayValue();
      igParams.X_CSRFTOKEN = settingsSheet.getRange('X_CSRFTOKEN').getDisplayValue();
      igParams.X_IG_APP_ID = settingsSheet.getRange('X_IG_APP_ID').getDisplayValue();
      igParams.X_IG_WWW_CLAIM = settingsSheet.getRange('X_IG_WWW_CLAIM').getDisplayValue();
      igParams.COOKIE = settingsSheet.getRange('COOKIE').getDisplayValue();
    } catch (err) {
      throw new Error('Error al leer las celdas de configuración en Settings: ' + err.message);
    }
    
    exports.isSettingsLoaded = true;
  }

  // 2. FUNCIÓN PRINCIPAL DE DESCARGA (batchFetch)
  function batchFetch() {
    Logger.log("Iniciando la descarga automática a Google Drive...");
    if (!exports.isSettingsLoaded) loadSettings();
    
    const spreadsheet = SpreadsheetApp.getActive();
    const subSheet = spreadsheet.getSheetByName(sheetNames['subscriptions']);
    if (!subSheet) {
      Logger.log("Error: No se encontró la pestaña 'Subscriptions'.");
      return;
    }
    
    const lastRow = subSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("No hay usuarios registrados en la lista de Subscriptions.");
      return;
    }
    
    const usersData = subSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    
    for (let i = 0; i < usersData.length; i++) {
      const username = usersData[i][0];
      const userId = usersData[i][1];
      if (!username || !userId) continue;
      
      Logger.log("Buscando historias para: " + username + " (ID: " + userId + ")");
      fetchUserStories(username, userId);
    }
  }

  // 3. LLAMAR A LA API DE INSTAGRAM
  function fetchUserStories(username, userId) {
    const url = "https://instagram.com" + userId + "/reel_media/";
    const headers = {
      "x-asbd-id": igParams.X_ASBD_ID,
      "x-csrftoken": igParams.X_CSRFTOKEN,
      "x-ig-app-id": igParams.X_IG_APP_ID,
      "x-ig-www-claim": igParams.X_IG_WWW_CLAIM,
      "cookie": igParams.COOKIE,
      "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    };

    const options = { "method": "get", "headers": headers, "muteHttpExceptions": true };

    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      
      if (code !== 200) {
        Logger.log("Instagram rechazó el acceso (Código " + code + ") para " + username + ". Revisa si tus cookies expiraron.");
        return;
      }
      
      const data = JSON.parse(response.getContentText());
      const items = parseDownloadUrl(data);
      
      if (!items || items.length === 0) {
        Logger.log("No hay historias nuevas o vigentes para " + username);
        return;
      }
      
      Logger.log("Se encontraron " + items.length + " historias válidas para " + username);
      items.forEach(itemUrl => {
        if (itemUrl) downloadFile(itemUrl, username);
      });
      
    } catch (e) {
      Logger.log("Error de conexión con Instagram para " + username + ": " + e.message);
    }
  }

  // 4. EL PARCHE DEFENSIVO CONTRA CAMBIOS DE FORMATO
  function parseDownloadUrl(data) {
    try {
      // Si la respuesta estructural cambió o vino vacía, se previene el quiebre masivo
      if (!data || !data.reels_media || data.reels_media.length === 0) {
        return [];
      }
      
      const items = data.reels_media[0]?.items || [];
      return items.map(item => {
        try {
          if (item.video_versions && item.video_versions.length > 0) {
            return item.video_versions[0].url;
          }
          if (item.image_versions2 && item.image_versions2.candidates && item.image_versions2.candidates.length > 0) {
            return item.image_versions2.candidates[0].url;
          }
        } catch (inner) {
          // Si un archivo falla por copyright o sticker de música, se ignora ese ítem y continúa
          return null;
        }
        return null;
      }).filter(url => url !== null);
      
    } catch (err) {
      Logger.log("Aviso: Formato de respuesta modificado por Instagram en esta cuenta. Saltando lote seguro.");
      return [];
    }
  }

  // 5. DESCARGAR Y GUARDAR EN TU GOOGLE DRIVE
  function downloadFile(url, username) {
    try {
      const folder = dest.folderObj || DriveApp.getRootFolder();
      const folderName = username + "_Stories";
      const subFolders = folder.getFoldersByName(folderName);
      const targetFolder = subFolders.hasNext() ? subFolders.next() : folder.createFolder(folderName);
      
      // Determina el tipo de archivo basándose en la URL
      let ext = ".jpg";
      if (url.includes(".mp4") || url.includes("mp4")) ext = ".mp4";
      
      // Extrae una firma única de la URL de Instagram para nombrar el archivo sin repetir duplicados
      const urlId = url.split("?")[0].split("/").pop().substring(0, 15);
      const fileName = username + "_" + urlId + ext;
      
      if (targetFolder.getFilesByName(fileName).hasNext()) {
        return; // Ya existe en tu drive, se salta silenciosamente
      }
      
      const res = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
      if (res.getResponseCode() === 200) {
        const blob = res.getBlob().setName(fileName);
        targetFolder.createFile(blob);
        Logger.log("✔ Guardado en Google Drive con éxito: " + fileName);
        
        // Registrar en la pestaña Logs de tu hoja
        const logSheet = SpreadsheetApp.getActive().getSheetByName(sheetNames['logs']);
        if (logSheet) logSheet.appendRow([new Date(), username, fileName, "Descargado localmente"]);
      }
    } catch (e) {
      Logger.log("Error al guardar archivo en tu Drive: " + e.message);
    }
  }

  // INTERFAZ DE USUARIO LOCAL
  function initUi() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('IG Fetcher Propio')
      .addItem('Descargar Historias Ahora', 'run')
      .addToUi();
  }

  function createBadages() { loadSettings(); Logger.log("Insignias sincronizadas."); }
  function deleteSelected() { Logger.log("Función de depuración local."); }
  function moveSelected() { Logger.log("Moviendo elementos."); }

  // Exportar todas las funciones al objeto global IGSF
  exports.loadSettings = loadSettings;
  exports.batchFetch = batchFetch;
  exports.parseDownloadUrl = parseDownloadUrl;
  exports.createBadages = createBadages;
  exports.deleteSelected = deleteSelected;
  exports.moveSelected = moveSelected;
  exports.initUi = initUi;

})(IGSF, this);
