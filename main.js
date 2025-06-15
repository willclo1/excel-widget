const { app, BrowserWindow, ipcMain, dialog } = require('electron')
const path = require('path')
const fs = require('fs')
const AsposeCells = require("aspose.cells.node")

// Template definitions optimized for 500x500
const templates = {
  modern: {
    name: 'Modern',
    css: generateModernCSS()
  },
  minimal: {
    name: 'Minimal',
    css: generateMinimalCSS()
  },
  colorful: {
    name: 'Colorful',
    css: generateColorfulCSS()
  },
  dark: {
    name: 'Dark Mode',
    css: generateDarkCSS()
  }
};

async function generateHtml(inputFilePath = null, templateName = 'modern') {
  try {
    let inputPath;
    if (inputFilePath) {
      inputPath = inputFilePath;
    } else {
      const isPackaged = app.isPackaged;
      if (isPackaged) {
        inputPath = path.join(process.resourcesPath, "weekly_schedule_12hr_template.xlsx");
      } else {
        inputPath = path.join(__dirname, "weekly_schedule_12hr_template.xlsx");
      }
    }

    const outputPath = path.join(__dirname, "output.html");
    console.log("Processing Excel file at:", inputPath);
    console.log("Using template:", templateName);

    if (!fs.existsSync(inputPath)) {
      throw new Error(`Excel file not found at: ${inputPath}`);
    }

    const workbook = new AsposeCells.Workbook(inputPath);
    const options = new AsposeCells.HtmlSaveOptions();
    options.saveAsSingleFile = true;
    workbook.save(outputPath, options);

    await addTemplateCSS(outputPath, templateName);
    console.log(`✅ output.html generated from: ${inputPath} with ${templateName} template`);
    return outputPath;
  } catch (err) {
    console.error("❌ Failed to generate output.html:", err);
    throw err;
  }
}

async function addTemplateCSS(htmlPath, templateName) {
  try {
    let html = fs.readFileSync(htmlPath, 'utf8');
    const template = templates[templateName] || templates.modern;
    const templateCSS = template.css;

    if (html.includes('</head>')) {
      html = html.replace('</head>', templateCSS + '\n</head>');
    } else {
      html = html.replace('</body>', templateCSS + '\n</body>');
    }

    // Compact back button for 500x500 widget
    const backButton = `
      <button id="backBtn" style="
        position: fixed;
        top: 8px;
        left: 8px;
        padding: 6px 12px;
        background: rgba(79, 70, 229, 0.9);
        color: white;
        border: none;
        border-radius: 6px;
        font-family: inherit;
        font-size: 11px;
        cursor: pointer;
        z-index: 999;
        backdrop-filter: blur(10px);
        border: 1px solid rgba(255, 255, 255, 0.2);
        transition: all 0.2s ease;
      " onmouseover="this.style.background='rgba(79, 70, 229, 1)'" 
         onmouseout="this.style.background='rgba(79, 70, 229, 0.9)'">
        ← Back
      </button>
      <script>
        document.getElementById('backBtn')
          .addEventListener('click', () => history.back());
      </script>
    `;

    html = html.replace(
      /<body([^>]*)>/,
      `<body$1>\n${backButton}`
    );

    fs.writeFileSync(htmlPath, html, 'utf8');
    console.log(`✅ ${template.name} template CSS added!`);
  } catch (err) {
    console.error("❌ Failed to add template CSS:", err);
  }
}

// Modern template optimized for 500x500
function generateModernCSS() {
  return `
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
      
      * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
      }
      
      html, body {
        height: 500px !important;
        width: 500px !important;
        overflow: hidden !important;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        font-family: 'Inter', sans-serif !important;
        padding: 12px;
        font-size: 11px !important;
      }
      
      #section {
        height: 476px !important;
        width: 476px !important;
        overflow: auto !important;
        background: rgba(255, 255, 255, 0.95) !important;
        backdrop-filter: blur(20px) !important;
        border-radius: 16px !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
        box-shadow: 0 8px 32px rgba(31, 38, 135, 0.37) !important;
        padding: 16px !important;
        margin-top: 8px;
      }
      
      table {
        width: 100% !important;
        border-collapse: separate !important;
        border-spacing: 0 !important;
        background: white !important;
        border-radius: 12px !important;
        overflow: hidden !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1) !important;
        font-size: 10px !important;
      }
      
      td, th {
        padding: 8px 10px !important;
        text-align: left !important;
        vertical-align: middle !important;
        font-size: 10px !important;
        border: none !important;
        line-height: 1.3 !important;
      }
      
      tr:first-child td,
      tr:first-child th {
        background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%) !important;
        color: white !important;
        font-weight: 600 !important;
        font-size: 11px !important;
        padding: 10px !important;
      }
      
      td:first-child,
      th:first-child {
        background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%) !important;
        color: #334155 !important;
        font-weight: 600 !important;
        font-size: 10px !important;
        border-right: 1px solid rgba(148, 163, 184, 0.2) !important;
        min-width: 60px;
      }
      
      tr:nth-child(even) td:not(:first-child) {
        background: #fafbff !important;
      }
      
      tr:nth-child(odd) td:not(:first-child) {
        background: white !important;
      }
      
      tr:hover td:not(:first-child) {
        background: linear-gradient(135deg, #f0f4ff 0%, #e0e7ff 100%) !important;
      }
      
      .x22, div[style*="color:red"] {
        display: none !important;
      }
      
      ::-webkit-scrollbar {
        width: 6px;
        height: 6px;
      }
      
      ::-webkit-scrollbar-track {
        background: rgba(248, 250, 252, 0.5);
        border-radius: 3px;
      }
      
      ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, #6366f1, #8b5cf6);
        border-radius: 3px;
      }
    </style>`;
}

// Minimal template optimized for 500x500
function generateMinimalCSS() {
  return `
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
      
      * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
      }
      
      html, body {
        height: 500px !important;
        width: 500px !important;
        overflow: hidden !important;
        background: #fcfcfc !important;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        padding: 12px;
      }
      
      #section {
        height: 476px !important;
        width: 476px !important;
        overflow: auto !important;
        background: white !important;
        border-radius: 12px !important;
        border: 1px solid #e5e7eb !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1) !important;
        padding: 16px !important;
        margin-top: 8px;
      }
      
      table {
        width: 100% !important;
        border-collapse: collapse !important;
        background: white !important;
        font-size: 10px !important;
      }
      
      td, th {
        padding: 8px 10px !important;
        text-align: left !important;
        vertical-align: middle !important;
        font-size: 10px !important;
        border-bottom: 1px solid #f1f5f9 !important;
        color: #374151 !important;
        line-height: 1.3 !important;
      }
      
      tr:first-child td,
      tr:first-child th {
        background: #f9fafb !important;
        color: #111827 !important;
        font-weight: 600 !important;
        font-size: 11px !important;
        border-bottom: 2px solid #e5e7eb !important;
        padding: 10px !important;
      }
      
      td:first-child,
      th:first-child {
        color: #111827 !important;
        font-weight: 500 !important;
        background: #fafbfc !important;
        border-right: 1px solid #f1f5f9 !important;
        min-width: 60px;
      }
      
      tr:hover td:not(:first-child) {
        background: #f8fafc !important;
      }
      
      tr:last-child td {
        border-bottom: none !important;
      }
      
      .x22, div[style*="color:red"] {
        display: none !important;
      }
      
      ::-webkit-scrollbar {
        width: 4px;
        height: 4px;
      }
      
      ::-webkit-scrollbar-track {
        background: transparent;
      }
      
      ::-webkit-scrollbar-thumb {
        background: #d1d5db;
        border-radius: 2px;
      }
    </style>`;
}

// Colorful template optimized for 500x500
function generateColorfulCSS() {
  return `
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
      
      * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
      }
      
      html, body {
        height: 500px !important;
        width: 500px !important;
        overflow: hidden !important;
        background: linear-gradient(135deg, #ff9a9e 0%, #fecfef 25%, #a8edea 75%, #fed6e3 100%) !important;
        font-family: 'Poppins', sans-serif !important;
        padding: 12px;
      }
      
      #section {
        height: 476px !important;
        width: 476px !important;
        overflow: auto !important;
        background: rgba(255, 255, 255, 0.9) !important;
        backdrop-filter: blur(25px) !important;
        border-radius: 16px !important;
        border: 1px solid rgba(255, 255, 255, 0.3) !important;
        box-shadow: 0 8px 25px rgba(0, 0, 0, 0.1) !important;
        padding: 16px !important;
        margin-top: 8px;
      }
      
      table {
        width: 100% !important;
        border-collapse: separate !important;
        border-spacing: 2px !important;
        background: transparent !important;
        border-radius: 12px !important;
        font-size: 10px !important;
      }
      
      td, th {
        padding: 8px 10px !important;
        text-align: left !important;
        vertical-align: middle !important;
        font-size: 10px !important;
        border: none !important;
        border-radius: 8px !important;
        line-height: 1.3 !important;
      }
      
      tr:first-child td,
      tr:first-child th {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        color: white !important;
        font-weight: 600 !important;
        font-size: 11px !important;
        box-shadow: 0 2px 8px rgba(102, 126, 234, 0.3) !important;
        padding: 10px !important;
      }
      
      td:first-child,
      th:first-child {
        background: linear-gradient(135deg, #a8edea 0%, #fed6e3 100%) !important;
        color: #2d3748 !important;
        font-weight: 600 !important;
        min-width: 60px;
      }
      
      tr:nth-child(even) td:not(:first-child) {
        background: linear-gradient(135deg, #ffecd2 0%, #fcb69f 100%) !important;
        color: #2d3748 !important;
      }
      
      tr:nth-child(odd) td:not(:first-child) {
        background: linear-gradient(135deg, #a8edea 0%, #fed6e3 100%) !important;
        color: #2d3748 !important;
      }
      
      tr:hover td:not(:first-child) {
        background: linear-gradient(135deg, #ff9a9e 0%, #fad0c4 100%) !important;
        color: white !important;
        transform: scale(1.01) !important;
      }
      
      .x22, div[style*="color:red"] {
        display: none !important;
      }
      
      ::-webkit-scrollbar {
        width: 6px;
        height: 6px;
      }
      
      ::-webkit-scrollbar-track {
        background: rgba(255, 255, 255, 0.3);
        border-radius: 3px;
      }
      
      ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, #667eea, #764ba2);
        border-radius: 3px;
      }
    </style>`;
}

// Dark template optimized for 500x500
function generateDarkCSS() {
  return `
    <style>
      @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@300;400;500;600;700&display=swap');
      
      * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
      }
      
      html, body {
        height: 500px !important;
        width: 500px !important;
        overflow: hidden !important;
        background: linear-gradient(135deg, #0c0c0c 0%, #1a1a2e 100%) !important;
        font-family: 'JetBrains Mono', monospace !important;
        padding: 12px;
      }
      
      #section {
        height: 476px !important;
        width: 476px !important;
        overflow: auto !important;
        background: rgba(15, 15, 15, 0.95) !important;
        backdrop-filter: blur(20px) !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        border-radius: 16px !important;
        box-shadow: 0 8px 25px rgba(0, 0, 0, 0.5) !important;
        padding: 16px !important;
        margin-top: 8px;
      }
      
      table {
        width: 100% !important;
        border-collapse: separate !important;
        border-spacing: 1px !important;
        background: #111111 !important;
        border-radius: 12px !important;
        overflow: hidden !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3) !important;
        font-size: 10px !important;
      }
      
      td, th {
        padding: 8px 10px !important;
        text-align: left !important;
        vertical-align: middle !important;
        font-size: 10px !important;
        border: none !important;
        color: #e2e8f0 !important;
        line-height: 1.3 !important;
      }
      
      tr:first-child td,
      tr:first-child th {
        background: linear-gradient(135deg, #1e3a8a 0%, #3730a3 100%) !important;
        color: #e0e7ff !important;
        font-weight: 600 !important;
        font-size: 11px !important;
        padding: 10px !important;
      }
      
      td:first-child,
      th:first-child {
        background: linear-gradient(135deg, #374151 0%, #4b5563 100%) !important;
        color: #f1f5f9 !important;
        font-weight: 500 !important;
        border-right: 1px solid rgba(75, 85, 99, 0.3) !important;
        min-width: 60px;
      }
      
      tr:nth-child(even) td:not(:first-child) {
        background: #1a1a1a !important;
      }
      
      tr:nth-child(odd) td:not(:first-child) {
        background: #111111 !important;
      }
      
      tr:hover td:not(:first-child) {
        background: linear-gradient(135deg, #1e40af 0%, #1d4ed8 100%) !important;
        color: #e0e7ff !important;
      }
      
      .x22, div[style*="color:red"] {
        display: none !important;
      }
      
      ::-webkit-scrollbar {
        width: 6px;
        height: 6px;
      }
      
      ::-webkit-scrollbar-track {
        background: rgba(55, 65, 81, 0.3);
        border-radius: 3px;
      }
      
      ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, #1e40af, #3730a3);
        border-radius: 3px;
      }
    </style>`;
}

const createWindow = async () => {
  const win = new BrowserWindow({
    width: 500,
    height: 500,
    frame: false,
    resizable: false,
    x: 1000,
    y: 0,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      icon: path.join(__dirname, 'resources', 'myIcon.icns'),
      contextIsolation: true
    }
  });

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
    return {
      canceled: false,
      filePath: filePath,
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

ipcMain.handle('process-with-template', async (event, { filePath, template }) => {
  try {
    const outputPath = await generateHtml(filePath, template);
    return {
      success: true,
      outputPath: outputPath
    };
  } catch (error) {
    console.error("Error processing file with template:", error);
    return {
      success: false,
      error: error.message
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