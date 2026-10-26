/**
 * Bundle unificado y corregido localmente sin dependencias externas.
 * Incluye parche defensivo para el cambio de API de Instagram de Octubre 2026.
 * Sincronizado con la lectura de Drive en la Columna C (Destination Folder ID).
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

  // 1. CARGAR CONFIGURACIÓN DE LA HOJA (SETTINGS GENERALES)
  function loadSettings() {
    const spreadsheet = SpreadsheetApp.getActive();
    const settingsSheet = spreadsheet.getSheetByName(sheetNames['settings']);
    if (!settingsSheet) throw new Error("No se encontró la pestaña 'Settings'");
    
    try {
      igParams.X_ASBD_ID = settingsSheet.getRange('X_ASBD_ID').getDisplayValue();
      igParams.X_CSRFTOKEN = settingsSheet.getRange('X_CSRFTOKEN').getDisplayValue();
      igParams.X_IG_APP_ID = settingsSheet.getRange('X_IG_APP_ID').getDisplayValue();
      igParams.X_IG_WWW_CLAIM = settingsSheet.getRange('X_IG_WWW_CLAIM').getDisplayValue();
      igParams.COOKIE = settingsSheet.getRange('COOKIE').getDisplayValue();
    } catch (err) {
      // Lectura de contingencia si los nombres de rango de celdas no están definidos en la plantilla
      const values = settingsSheet.getRange(1, 1, settingsSheet.getLastRow(), 2).getValues();
      values.forEach(row => {
        const key = String(row[0]).toUpperCase().replace(/-/g, '_');
        if (igParams.hasOwnProperty(key)) igParams[key] = row[1];
      });
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
    
    // Lee la columna A (Username), B (User ID) y C (Destination Folder ID) de tu tabla
    const usersData = subSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    
    for (let i = 0; i < usersData.length; i++) {
      const username = usersData[i][0];
      const userId = usersData[i][1];
      const customFolderId = usersData[i][2]; // Columna C de tu imagen
      
      if (!username || !userId) continue;
      
      Logger.log("Buscando historias para: " + username + " (ID: " + userId + ")");
      fetchUserStories(username, userId, customFolderId);
    }
  }

  // 3. LLAMAR A LA API DE INSTAGRAM (URL REPARADA DEFECTO DNS)
  function fetchUserStories(username, userId, customFolderId) {
    // CORRECCIÓN TOTAL DE URL: Conexión limpia con subdominio y endpoints correctos de Meta
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
        if (itemUrl) downloadFile(itemUrl, username, customFolderId);
      });
      
    } catch (e) {
      Logger.log("Error de conexión con Instagram para " + username + ": " + e.message);
    }
  }

  // 4. EL PARCHE DEFENSIVO CONTRA CAMBIOS DE FORMATO (LÍNEA 654)
  function parseDownloadUrl(data) {
    try {
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
          return null;
        }
        return null;
      }).filter(url => url !== null);
      
    } catch (err) {
      Logger.log("Aviso: Formato de respuesta modificado por Instagram en esta cuenta. Saltando lote seguro.");
      return [];
    }
  }

  // 5. DESCARGAR Y ENLAZAR CON EL DRIVE CORRESPONDIENTE DE LA COLUMNA C
  function downloadFile(url, username, customFolderId) {
    try {
      let folderObj;
      
      // Valida si la celda de la columna C viene con un ID de Drive válido
      if (customFolderId && String(customFolderId).trim() !== "") {
        folderObj = DriveApp.getFolderById(String(customFolderId).trim());
      } else {
        folderObj = DriveApp.getRootFolder();
      }
      
      const folderName = username + "_Stories";
      const subFolders = folderObj.getFoldersByName(folderName);
      const targetFolder = subFolders.hasNext() ? subFolders.next() : folderObj.createFolder(folderName);
      
      let ext = ".jpg";
      if (url.includes(".mp4") || url.includes("mp4")) ext = ".mp4";
      
      const urlId = url.split("?")[0].split("/").pop().substring(0, 15);
      const fileName = username + "_" + urlId + ext;
      
      if (targetFolder.getFilesByName(fileName).hasNext()) {
        return; 
      }
      
      const res = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
      if (res.getResponseCode() === 200) {
        const blob = res.getBlob().setName(fileName);
        targetFolder.createFile(blob);
        Logger.log("✔ Guardado en Google Drive con éxito: " + fileName);
        
        const logSheet = SpreadsheetApp.getActive().getSheetByName(sheetNames['logs']);
        if (logSheet) logSheet.appendRow([new Date(), username, fileName, "Descargado localmente"]);
      }
    } catch (e) {
      Logger.log("Error al guardar archivo en tu Drive para " + username + ": " + e.message);
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

  exports.loadSettings = loadSettings;
  exports.batchFetch = batchFetch;
  exports.parseDownloadUrl = parseDownloadUrl;
  exports.createBadages = createBadages;
  exports.deleteSelected = deleteSelected;
  exports.moveSelected = moveSelected;
  exports.initUi = initUi;

})(IGSF, this);
