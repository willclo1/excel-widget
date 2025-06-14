const { app, BrowserWindow, ipcMain, dialog } = require('electron')
const path = require('path')
const fs = require('fs')
const AsposeCells = require("aspose.cells.node")

// Template definitions
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
    console.log("Using template:", templateName);
    
    if (!fs.existsSync(inputPath)) {
      throw new Error(`Excel file not found at: ${inputPath}`);
    }

    const workbook = new AsposeCells.Workbook(inputPath);
    const options = new AsposeCells.HtmlSaveOptions();
    options.saveAsSingleFile = true;

    workbook.save(outputPath, options);
    
    // Add template-specific CSS to the generated HTML
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
    
    // Insert the CSS before the closing </head> tag or before </body> if no head
    if (html.includes('</head>')) {
      html = html.replace('</head>', templateCSS + '\n</head>');
    } else {
      html = html.replace('</body>', templateCSS + '\n</body>');
    }

    const backButton = `
      <button id="backBtn" style="
        position: fixed;
        top: 24px;
        left: 24px;
        padding: 0.5rem 1rem;
        background: #4f46e5;
        color: #fff;
        border: none;
        border-radius: 4px;
        font-family: inherit;
        cursor: pointer;
        z-index: 999;
      ">
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
function generateModernCSS() {
  return `
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@300;400;500;600;700&display=swap');
      
      * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
      }
      
      html, body {
        height: 100vh !important;
        width: 100vw !important;
        overflow: hidden !important;
        background: linear-gradient(145deg, #e5e7eb, #f3f4f6) !important;
        font-family: 'Manrope', sans-serif !important;
        padding: 24px;
      }
      
      #section {
        height: calc(100vh - 48px) !important;
        width: calc(100vw - 48px) !important;
        overflow: auto !important;
        background: rgba(255, 255, 255, 0.98) !important;
        backdrop-filter: blur(12px) !important;
        border-radius: 24px !important;
        box-shadow: 0 12px 48px rgba(0, 0, 0, 0.08) !important;
        padding: 32px !important;
      }
      
      table {
        width: 100% !important;
        border-collapse: separate !important;
        border-spacing: 0 !important;
        background: white !important;
        border-radius: 16px !important;
        overflow: hidden !important;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.06) !important;
      }
      
      td, th {
        padding: 16px 24px !important;
        text-align: left !important;
        vertical-align: middle !important;
        font-size: 15px !important;
        border: none !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
      }
      
      tr:first-child td {
        background: linear-gradient(90deg, #3b82f6, #60a5fa) !important;
        color: black !important;
        font-weight: bold !important;
        font-size: 16px !important;
        letter-spacing: 0.5px !important;
        
      }
      
      td:first-child {
        background: #f1f5f9 !important;
        color: #1e293b !important;
        font-weight: bold !important;
        font-size: 14px !important;
        min-width: 140px !important;
      }

      /* cover both td- and th-based headers */
      tr:first-child td,
      tr:first-child th,
      td:first-child,
      th:first-child {
        font-weight: bold !important;
      }
      
      tr:nth-child(even) td:not(:first-child) {
        background: #f9fafb !important;
      }
      
      tr:nth-child(odd) td:not(:first-child) {
        background: white !important;
      }
      
      tr:hover td:not(:first-child) {
        background: #eff6ff !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.15) !important;
      }
      
      .x22, div[style*="color:red"] {
        display: none !important;
      }
      
      ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
      }
      
      ::-webkit-scrollbar-track {
        background: rgba(241, 245, 249, 0.5);
        border-radius: 4px;
      }
      
      ::-webkit-scrollbar-thumb {
        background: linear-gradient(90deg, #60a5fa, #3b82f6);
        border-radius: 4px;
      }
    </style>`;
}

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
        height: 100vh !important;
        width: 100vw !important;
        overflow: hidden !important;
        background: #fafafa !important;
        font-family: 'Inter', sans-serif !important;
        padding: 20px;
      }
      
      #section {
        height: calc(100vh - 40px) !important;
        width: calc(100vw - 40px) !important;
        overflow: auto !important;
        background: white !important;
        border-radius: 16px !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05) !important;
        padding: 24px !important;
      }
      
      table {
        width: 100% !important;
        border-collapse: collapse !important;
        background: white !important;
      }
      
      td, th {
        padding: 14px 20px !important;
        text-align: left !important;
        vertical-align: middle !important;
        font-size: 14px !important;
        border-bottom: 1px solid #f1f5f9 !important;
        color: #1e293b !important;
      }
      
      tr:first-child td {
        background: #f8fafc !important;
        color: #1e293b !important;
        font-weight: bold !important;
        font-size: 15px !important;
        border-bottom: 2px solid #e5e7eb !important;
      }
      
      td:first-child {
        color: #374151 !important;
        font-weight: bold !important;
        font-size: 14px !important;
        background: #f9fafb !important;
      }
        /* cover both td- and th-based headers */
      tr:first-child td,
      tr:first-child th,
      td:first-child,
      th:first-child {
        font-weight: bold !important;
      }
      
      tr:hover td:not(:first-child) {
        background: #f1f5f9 !important;
      }
      
      .x22, div[style*="color:red"] {
        display: none !important;
      }
    </style>`;
}

function generateColorfulCSS() {
  return `
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600&display=swap');
      
      * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
      }
      
      html, body {
        height: 100vh !important;
        width: 100vw !important;
        overflow: hidden !important;
        background: linear-gradient(135deg, #f9a8d4, #a5b4fc, #6ee7b7) !important;
        font-family: 'Outfit', sans-serif !important;
        padding: 24px;
      }
      
      #section {
        height: calc(100vh - 48px) !important;
        width: calc(100vw - 48px) !important;
        overflow: auto !important;
        background: rgba(255, 255, 255, 0.95) !important;
        backdrop-filter: blur(10px) !important;
        border-radius: 20px !important;
        box-shadow: 0 12px 40px rgba(0, 0, 0, 0.1) !important;
        padding: 32px !important;
      }
      
      table {
        width: 100% !important;
        border-collapse: separate !important;
        border-spacing: 3px !important;
        background: transparent !important;
      }
      
      td, th {
        padding: 16px 20px !important;
        text-align: left !important;
        vertical-align: middle !important;
        font-size: 14px !important;
        border: none !important;
        border-radius: 8px !important;
        transition: all 0.3s ease !important;
      }
      
      tr:first-child td {
        background: linear-gradient(90deg, #ec4899, #f472b6) !important;
        color: black !important;
        font-weight: bold !important;
        font-size: 15px !important;
      }
      
      td:first-child {
        background: linear-gradient(90deg, #f43f5e, #fb7185) !important;
        color: black !important;
        font-weight: bold !important;
        font-size: 14px !important;
      }
        /* cover both td- and th-based headers */
      tr:first-child td,
      tr:first-child th,
      td:first-child,
      th:first-child {
        font-weight: bold !important;
      }
      
      
      tr:nth-child(even) td:not(:first-child):not([title]) {
        background: linear-gradient(90deg, #fff1f2, #ffe4e6) !important;
      }
      
      tr:nth-child(odd) td:not(:first-child):not([title]) {
        background: linear-gradient(90deg, #ecfdf5, #d1fae5) !important;
      }
      
      tr:hover td:not(:first-child) {
        transform: scale(1.02) !important;
        box-shadow: 0 6px 12px rgba(0, 0, 0, 0.08) !important;
      }
      
      .x22, div[style*="color:red"] {
        display: none !important;
      }
    </style>`;
}

function generateDarkCSS() {
  return `
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&display=swap');
      
      * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
      }
      
      html, body {
        height: 100vh !important;
        width: 100vw !important;
        overflow: hidden !important;
        background: linear-gradient(135deg, #1f2937, #374151) !important;
        font-family: 'Space Mono', monospace !important;
        padding: 24px;
      }
      
      #section {
        height: calc(100vh - 48px) !important;
        width: calc(100vw - 48px) !important;
        overflow: auto !important;
        background: rgba(17, 24, 39, 0.95) !important;
        backdrop-filter: blur(12px) !important;
        border: 1px solid rgba(59, 130, 246, 0.15) !important;
        border-radius: 20px !important;
        box-shadow: 0 12px 40px rgba(0, 0, 0, 0.2) !important;
        padding: 32px !important;
      }
      
      table {
        width: 100% !important;
        border-collapse: separate !important;
        border-spacing: 2px !important;
        background: #111827 !important;
        border-radius: 12px !important;
        overflow: hidden !important;
      }
      
      td, th {
        padding: 16px 20px !important;
        text-align: left !important;
        vertical-align: middle !important;
        font-size: 14px !important;
        border: none !important;
        color: #d1d5db !important;
        transition: all 0.3s ease !important;
      }
      
      tr:first-child td {
        background: linear-gradient(90deg, #2563eb, #3b82f6) !important;
        color: white !important;
        font-weight: bold !important;
        font-size: 15px !important;
      }
      
      td:first-child {
        background: linear-gradient(90deg, #1e3a8a, #1e40af) !important;
        color: white !important;
        font-weight: bold !important;
        font-size: 14px !important;
      }
      /* cover both td- and th-based headers */
      tr:first-child td,
      tr:first-child th,
      td:first-child,
      th:first-child {
        font-weight: bold !important;
      }
      tr:nth-child(even) td:not(:first-child) {
        background: #1f2937 !important;
      }
      
      tr:nth-child(odd) td:not(:first-child) {
        background: #111827 !important;
      }
      
      tr:hover td:not(:first-child) {
        background: linear-gradient(90deg, #4b5563, #6b7280) !important;
        color: white !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.2) !important;
      }
      
      .x22, div[style*="color:red"] {
        display: none !important;
      }
      
      ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
      }
      
      ::-webkit-scrollbar-track {
        background: rgba(31, 41, 55, 0.4);
        border-radius: 4px;
      }
      
      ::-webkit-scrollbar-thumb {
        background: linear-gradient(90deg, #2563eb, #3b82f6);
        border-radius: 4px;
      }
    </style>`;
}

const createWindow = async () => {
  const win = new BrowserWindow({
    width: 2000,
    height: 2000,
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
    // Generate HTML from the selected file with the chosen template
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