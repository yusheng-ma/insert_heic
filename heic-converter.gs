/**
 * HEIC to JPEG Converter for Google Sheets
 * Converts HEIC images from a Google Drive folder to JPEG format
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🖼️ HEIC Converter')
    .addItem('Convert Single HEIC to JPEG', 'convertHeicToJpeg')
    .addItem('Convert All & Insert to Sheet', 'convertAllAndInsert')
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
        cellB2.clear();
        
        // Set row height and column width for better display
        sheet.setRowHeight(2, 300);
        sheet.setColumnWidth(2, 300);
        
        // Insert image with IMAGE formula
        cellB2.setFormula(`=IMAGE("${imageUrl}", 1)`);
        
        ui.alert(
          'Conversion Successful! ✅',
          `File converted and embedded into cell B2:\n\n` +
          `Original: ${result.originalName}\n` +
          `New: ${result.newName}\n` +
          `Location: ${sourceFolder.getName()}`,
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
 * Convert all HEIC files and insert them into the sheet
 * Starting from B2, 6 images per row
 */
function convertAllAndInsert() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Prompt user for folder link
  const folderResponse = ui.prompt(
    'Convert All HEIC Files',
    'Paste the link to the folder containing HEIC files (~40 files):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (folderResponse.getSelectedButton() != ui.Button.OK) {
    return;
  }
  
  const folderLink = folderResponse.getResponseText().trim();
  
  try {
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
    
    // Sort files by name (ascending)
    heicFiles.sort((a, b) => {
      return a.getName().localeCompare(b.getName(), undefined, { numeric: true, sensitivity: 'base' });
    });
    
    const confirmResponse = ui.alert(
      'Confirm Batch Conversion',
      `Found ${heicFiles.length} HEIC files.\n\n` +
      `This will:\n` +
      `1. Convert all files to JPEG\n` +
      `2. Delete original HEIC files\n` +
      `3. Insert images in grid (6 per row) starting at B2\n\n` +
      `Continue?`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse != ui.Button.YES) {
      return;
    }
    
    // Show progress
    ui.alert('Processing...', `Converting ${heicFiles.length} files. This may take a few minutes...`, ui.ButtonSet.OK);
    
    const convertedFiles = [];
    let successCount = 0;
    let failCount = 0;
    
    // Convert all files
    heicFiles.forEach((file, index) => {
      try {
        const result = convertSingleHeicToJpeg(file, sourceFolder);
        
        if (result.success) {
          successCount++;
          convertedFiles.push({
            fileId: result.newFileId,
            name: result.newName
          });
          
          // Delete original HEIC file
          file.setTrashed(true);
          
        } else {
          failCount++;
          console.error(`Failed to convert: ${file.getName()}`);
        }
      } catch (e) {
        failCount++;
        console.error(`Error processing ${file.getName()}:`, e);
      }
    });
    
    if (convertedFiles.length === 0) {
      ui.alert('Error', 'No files were successfully converted.', ui.ButtonSet.OK);
      return;
    }
    
    // Insert images into sheet
    const IMAGES_PER_ROW = 6;
    const START_ROW = 2;  // Row 2
    const START_COL = 2;  // Column B
    const CELL_SIZE = 200; // pixels
    
    // Set all row heights and column widths first
    const totalRows = Math.ceil(convertedFiles.length / IMAGES_PER_ROW);
    for (let i = 0; i < totalRows; i++) {
      sheet.setRowHeight(START_ROW + i, CELL_SIZE);
    }
    for (let i = 0; i < IMAGES_PER_ROW; i++) {
      sheet.setColumnWidth(START_COL + i, CELL_SIZE);
    }
    
    // Insert each image
    convertedFiles.forEach((file, index) => {
      try {
        const newFile = DriveApp.getFileById(file.fileId);
        
        // Make file accessible
        newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        // Calculate position
        const rowOffset = Math.floor(index / IMAGES_PER_ROW);
        const colOffset = index % IMAGES_PER_ROW;
        const row = START_ROW + rowOffset;
        const col = START_COL + colOffset;
        
        // Get image URL
        const imageUrl = `https://drive.google.com/uc?export=view&id=${file.fileId}`;
        
        // Insert image with IMAGE formula
        const cell = sheet.getRange(row, col);
        cell.setFormula(`=IMAGE("${imageUrl}", 1)`);
        
      } catch (e) {
        console.error(`Error inserting image ${file.name}:`, e);
      }
    });
    
    SpreadsheetApp.flush();
    
    ui.alert(
      'Batch Conversion Complete! 🎉',
      `✅ Successfully converted: ${successCount}\n` +
      `❌ Failed: ${failCount}\n` +
      `📸 Images inserted: ${convertedFiles.length}\n\n` +
      `Layout: ${IMAGES_PER_ROW} images per row, starting at B2\n` +
      `Original HEIC files have been deleted.`,
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    ui.alert('Error', `An error occurred: ${e.toString()}`, ui.ButtonSet.OK);
    console.error(e);
  }
}

/**
 * Extract folder ID from various Google Drive folder URL formats
 */
function extractFolderId(link) {
  let match = link.match(/[-\w]{25,}/);
  if (match) {
    return match[0];
  }
  
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
    const url = `https://drive.google.com/thumbnail?id=${fileId}&sz=w1000`;
    const blob = UrlFetchApp.fetch(url, {
      headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() }
    }).getBlob().getAs("image/jpeg");
    
    const newName = name.split('.').slice(0, -1).join('.') + '.jpg';
    blob.setName(newName);
    
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
