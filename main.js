const { app, BrowserWindow,  ipcMain, dialog } = require('electron')
const path = require('path')
const fs = require('fs')
const AsposeCells = require("aspose.cells.node")

async function generateHtml(inputFilePath = null) {
  try {
    let inputPath;
    
    if (inputFilePath) {
      // Use the user-selected file
      inputPath = inputFilePath;
    } else {
      // Fall back to default template (existing logic)
      const isPackaged = app.isPackaged;
      if (isPackaged) {
        inputPath = path.join(process.resourcesPath, "weekly_schedule_12hr_template.xlsx");
      } else {
        inputPath = path.join(__dirname, "weekly_schedule_12hr_template.xlsx");
      }
    }
    
    const outputPath = path.join(__dirname, "output.html");
    
    console.log("Processing Excel file at:", inputPath);
    
    if (!fs.existsSync(inputPath)) {
      throw new Error(`Excel file not found at: ${inputPath}`);
    }

    const workbook = new AsposeCells.Workbook(inputPath);
    const options = new AsposeCells.HtmlSaveOptions();
    options.saveAsSingleFile = true;

    workbook.save(outputPath, options);
    
    // Add beautiful CSS to the generated HTML
    await addPrettyCss(outputPath);
    
    console.log("✅ output.html generated from:", inputPath);
    return outputPath;
  } catch (err) {
    console.error("❌ Failed to generate output.html:", err);
    throw err;
  }
}

async function addPrettyCss(htmlPath) {
  try {
    let html = fs.readFileSync(htmlPath, 'utf8');
    
    const prettyCss = `
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
      
      * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
      }
      
      html, body {
        height: 100vh !important;
        width: 100vw !important;
        overflow: hidden !important;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        font-family: 'Inter', system-ui, -apple-system, sans-serif !important;
        padding: 20px;
      }
      
      #section {
        height: calc(100vh - 40px) !important;
        width: calc(100vw - 40px) !important;
        overflow: auto !important;
        background: rgba(255, 255, 255, 0.95) !important;
        backdrop-filter: blur(20px) !important;
        border-radius: 24px !important;
        box-shadow: 0 25px 50px rgba(0, 0, 0, 0.25) !important;
        padding: 30px !important;
      }
      
      table {
        width: 100% !important;
        height: 100% !important;
        border-collapse: separate !important;
        border-spacing: 0 !important;
        background: linear-gradient(145deg, #ffffff, #f8fafc) !important;
        border-radius: 16px !important;
        overflow: hidden !important;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.1) !important;
      }
      
      td, th {
        padding: 16px 20px !important;
        text-align: center !important;
        vertical-align: middle !important;
        font-weight: 500 !important;
        font-size: 14px !important;
        border: none !important;
        position: relative !important;
      }
      
      /* Header row styling */
      tr:first-child td {
        background: linear-gradient(135deg, #4f46e5, #7c3aed) !important;
        color: white !important;
        font-weight: 700 !important;
        font-size: 16px !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
        text-shadow: 0 2px 4px rgba(0, 0, 0, 0.3) !important;
      }
      
      /* Time column styling */
      td:first-child {
        background: linear-gradient(135deg, #1e293b, #334155) !important;
        color: white !important;
        font-weight: 600 !important;
        font-size: 13px !important;
        min-width: 100px !important;
      }
      
      /* Alternating row colors */
      tr:nth-child(even) td:not(:first-child) {
        background: linear-gradient(135deg, #f1f5f9, #e2e8f0) !important;
      }
      
      tr:nth-child(odd) td:not(:first-child) {
        background: linear-gradient(135deg, #ffffff, #f8fafc) !important;
      }
      
      /* Cell hover effects */
      td:not(:first-child) {
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
        cursor: pointer !important;
      }
      
      tr:hover td:not(:first-child) {
        background: linear-gradient(135deg, #ddd6fe, #c4b5fd) !important;
        transform: scale(1.02) !important;
        box-shadow: 0 8px 25px rgba(139, 92, 246, 0.3) !important;
        border-radius: 8px !important;
        z-index: 10 !important;
      }
      
      /* Activity-based coloring */
      td:contains("Work") {
        background: linear-gradient(135deg, #fef3c7, #fde68a) !important;
        color: #92400e !important;
        font-weight: 600 !important;
      }
      
      td:contains("Interview Prep") {
        background: linear-gradient(135deg, #dbeafe, #bfdbfe) !important;
        color: #1e40af !important;
        font-weight: 600 !important;
      }
      
      td:contains("Research") {
        background: linear-gradient(135deg, #d1fae5, #a7f3d0) !important;
        color: #065f46 !important;
        font-weight: 600 !important;
      }
      
      td:contains("Relax") {
        background: linear-gradient(135deg, #fce7f3, #fbcfe8) !important;
        color: #be185d !important;
        font-weight: 600 !important;
      }
      
      td:contains("Dinner") {
        background: linear-gradient(135deg, #fed7d7, #feb2b2) !important;
        color: #c53030 !important;
        font-weight: 600 !important;
      }
      
      td:contains("Lunch") {
        background: linear-gradient(135deg, #fef5e7, #feebc8) !important;
        color: #c05621 !important;
        font-weight: 600 !important;
      }
      
      td:contains("Independent Project") {
        background: linear-gradient(135deg, #e0e7ff, #c7d2fe) !important;
        color: #3730a3 !important;
        font-weight: 600 !important;
      }
      
      td:contains("Wake Up") {
        background: linear-gradient(135deg, #fef7cd, #fef3c7) !important;
        color: #a16207 !important;
        font-weight: 600 !important;
      }
      
      /* Corner radius for first and last cells */
      tr:first-child td:first-child {
        border-top-left-radius: 16px !important;
      }
      
      tr:first-child td:last-child {
        border-top-right-radius: 16px !important;
      }
      
      tr:last-child td:first-child {
        border-bottom-left-radius: 16px !important;
      }
      
      tr:last-child td:last-child {
        border-bottom-right-radius: 16px !important;
      }
      
      /* Hide evaluation warning */
      .x22, div[style*="color:red"] {
        display: none !important;
      }
      
      /* Smooth animations */
      * {
        transition: all 0.3s ease !important;
      }
      
      /* Scrollbar styling */
      ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
      }
      
      ::-webkit-scrollbar-track {
        background: rgba(255, 255, 255, 0.1);
        border-radius: 4px;
      }
      
      ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, #667eea, #764ba2);
        border-radius: 4px;
      }
      
      ::-webkit-scrollbar-thumb:hover {
        background: linear-gradient(135deg, #5a67d8, #6b46c1);
      }
    </style>`;
    
    // Insert the CSS before the closing </head> tag or before </body> if no head
    if (html.includes('</head>')) {
      html = html.replace('</head>', prettyCss + '\n</head>');
    } else {
      html = html.replace('</body>', prettyCss + '\n</body>');
    }
    
    fs.writeFileSync(htmlPath, html, 'utf8');
    console.log("✅ Beautiful CSS added!");
  } catch (err) {
    console.error("❌ Failed to add CSS:", err);
  }
}

const createWindow = async () => {
  const win = new BrowserWindow({
    width: 400,
    height: 400,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      icon: path.join(__dirname, 'resources', 'myIcon.icns'),
      contextIsolation: true
    }
  });


  // Load the initial interface
  const indexPath = path.join(__dirname, "index.html");
  win.loadFile(indexPath);
}

app.whenReady().then(() => {
  createWindow()

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow()
    }
  })
})
ipcMain.handle('open-file', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    properties: ['openFile'],
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls'] },
      { name: 'CSV Files', extensions: ['csv'] },
      { name: 'All Files', extensions: ['*'] }
    ]
  });

  if (canceled || filePaths.length === 0) {
    return { canceled: true };
  }

  const filePath = filePaths[0];
  
  try {
    // Generate HTML from the selected file
    const outputPath = await generateHtml(filePath);
    
    return { 
      canceled: false, 
      filePath: filePath,
      outputPath: outputPath,
      success: true 
    };
  } catch (error) {
    console.error("Error processing file:", error);
    return { 
      canceled: false, 
      error: error.message,
      success: false 
    };
  }
});
app.setLoginItemSettings({
    openAtLogin: true,
    path: app.getPath('exe')
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})

ipcMain.handle('reload-content', async (event, outputPath) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  if (win && fs.existsSync(outputPath)) {
    await win.loadFile(outputPath);
    return { success: true };
  }
  return { success: false };
});