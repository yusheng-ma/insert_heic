/**
 * HEIC to JPEG Converter for Google Sheets
 * Converts HEIC images from a Google Drive folder to JPEG format
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🖼️ HEIC Converter')
    .addItem('Convert HEIC to JPEG', 'convertHeicToJpeg')
    .addItem('Batch Convert All', 'batchConvertAllHeic')
    .addToUi();
}

function convertHeicToJpeg() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Step 1: Prompt user for folder link
  const folderResponse = ui.prompt(
    'Enter Google Drive Folder Link',
    'Paste the link to the folder containing HEIC files:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (folderResponse.getSelectedButton() != ui.Button.OK) {
    return;
  }
  
  const folderLink = folderResponse.getResponseText().trim();
  
  try {
    // Extract folder ID from the link
    const folderId = extractFolderId(folderLink);
    if (!folderId) {
      ui.alert('Error', 'Invalid folder link. Please provide a valid Google Drive folder link.', ui.ButtonSet.OK);
      return;
    }
    
    // Get the folder
    const sourceFolder = DriveApp.getFolderById(folderId);
    
    // Get all HEIC files
    const heicFiles = getHeicFiles(sourceFolder);
    
    if (heicFiles.length === 0) {
      ui.alert('No HEIC Files Found', 'The folder does not contain any HEIC files.', ui.ButtonSet.OK);
      return;
    }
    
    // Show available files and let user choose
    const fileNames = heicFiles.map((f, i) => `${i + 1}. ${f.getName()}`).join('\n');
    const fileResponse = ui.prompt(
      `Found ${heicFiles.length} HEIC file(s)`,
      `Enter the number of the file to convert (1-${heicFiles.length}):\n\n${fileNames}`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (fileResponse.getSelectedButton() != ui.Button.OK) {
      return;
    }
    
    const fileIndex = parseInt(fileResponse.getResponseText().trim()) - 1;
    
    if (isNaN(fileIndex) || fileIndex < 0 || fileIndex >= heicFiles.length) {
      ui.alert('Error', 'Invalid file number selected.', ui.ButtonSet.OK);
      return;
    }
    
    // Convert the selected file
    const file = heicFiles[fileIndex];
    const result = convertSingleHeicToJpeg(file, sourceFolder);
    
    if (result.success) {
      // Insert the converted image into cell B2 using IMAGE formula
      try {
        const newFile = DriveApp.getFileById(result.newFileId);
        
        // Make the file accessible (anyone with link can view)
        newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        // Get the file URL for IMAGE function
        const fileId = result.newFileId;
        const imageUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;
        
        // Clear cell B2 and insert image using IMAGE formula
        const cellB2 = sheet.getRange('B2');
        cellB2.clear(); // Clear any existing content
        
        // Set row height for better display (optional)
        sheet.setRowHeight(2, 300);
        
        // Set column width for better display (optional)
        sheet.setColumnWidth(2, 300);
        
        // Insert image with IMAGE formula - this embeds the image IN the cell
        // Syntax: =IMAGE(url, [mode])
        // mode 1 = fit to cell, 2 = stretch to fit, 3 = original size, 4 = custom size
        cellB2.setFormula(`=IMAGE("${imageUrl}", 1)`);
        
        ui.alert(
          'Conversion Successful! ✅',
          `File converted and embedded into cell B2:\n\n` +
          `Original: ${result.originalName}\n` +
          `New: ${result.newName}\n` +
          `Location: ${sourceFolder.getName()}\n\n` +
          `Image is now embedded in the cell!`,
          ui.ButtonSet.OK
        );
      } catch (e) {
        ui.alert('Warning', `File converted successfully but failed to embed into B2: ${e.toString()}`, ui.ButtonSet.OK);
        console.error('Error embedding image:', e);
      }
    } else {
      ui.alert('Conversion Failed', `Error: ${result.error}`, ui.ButtonSet.OK);
    }
    
  } catch (e) {
    ui.alert('Error', `An error occurred: ${e.toString()}`, ui.ButtonSet.OK);
    console.error(e);
  }
}

/**
 * Extract folder ID from various Google Drive folder URL formats
 */
function extractFolderId(link) {
  // Format: https://drive.google.com/drive/folders/FOLDER_ID
  let match = link.match(/[-\w]{25,}/);
  if (match) {
    return match[0];
  }
  
  // If it's already just an ID
  if (/^[-\w]{25,}$/.test(link)) {
    return link;
  }
  
  return null;
}

/**
 * Get all HEIC files from a folder
 */
function getHeicFiles(folder) {
  const heicFiles = [];
  const files = folder.getFiles();
  
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName().toLowerCase();
    const mimeType = file.getMimeType();
    
    if (name.includes('.heic') || name.includes('.heif') || mimeType === 'image/heic') {
      heicFiles.push(file);
    }
  }
  
  return heicFiles;
}

/**
 * Convert a single HEIC file to JPEG
 * Based on code from: https://stackoverflow.com/a/76624861
 */
function convertSingleHeicToJpeg(file, destFolder) {
  const fileId = file.getId();
  const name = file.getName();
  
  try {
    // Use Google Drive thumbnail API to convert HEIC to JPEG
    const url = `https://drive.google.com/thumbnail?id=${fileId}&sz=w1000`;
    const blob = UrlFetchApp.fetch(url, {
      headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() }
    }).getBlob().getAs("image/jpeg");
    
    // Create new filename
    const newName = name.split('.').slice(0, -1).join('.') + '.jpg';
    blob.setName(newName);
    
    // Save to destination folder
    const newFile = destFolder.createFile(blob);
    
    return {
      success: true,
      originalName: name,
      newName: newName,
      newFileId: newFile.getId()
    };
    
  } catch (e) {
    console.error(`Error converting ${name}:`, e);
    return {
      success: false,
      error: e.toString()
    };
  }
}

/**
 * Optional: Batch convert all HEIC files in a folder
 */
function batchConvertAllHeic() {
  const ui = SpreadsheetApp.getUi();
  
  const folderResponse = ui.prompt(
    'Batch Convert All HEIC Files',
    'Paste the link to the folder containing HEIC files:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (folderResponse.getSelectedButton() != ui.Button.OK) {
    return;
  }
  
  const folderLink = folderResponse.getResponseText().trim();
  const folderId = extractFolderId(folderLink);
  
  if (!folderId) {
    ui.alert('Error', 'Invalid folder link.', ui.ButtonSet.OK);
    return;
  }
  
  const sourceFolder = DriveApp.getFolderById(folderId);
  const heicFiles = getHeicFiles(sourceFolder);
  
  if (heicFiles.length === 0) {
    ui.alert('No HEIC Files Found', 'The folder does not contain any HEIC files.', ui.ButtonSet.OK);
    return;
  }
  
  const confirmResponse = ui.alert(
    'Confirm Batch Conversion',
    `Convert all ${heicFiles.length} HEIC files to JPEG?`,
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResponse != ui.Button.YES) {
    return;
  }
  
  let successCount = 0;
  let failCount = 0;
  
  heicFiles.forEach(file => {
    const result = convertSingleHeicToJpeg(file, sourceFolder);
    if (result.success) {
      successCount++;
    } else {
      failCount++;
    }
  });
  
  ui.alert(
    'Batch Conversion Complete',
    `✅ Successfully converted: ${successCount}\n` +
    `❌ Failed: ${failCount}`,
    ui.ButtonSet.OK
  );
}
