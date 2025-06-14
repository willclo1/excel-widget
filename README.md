<p align="center">
  <img src="logo.png" alt="Excel Tool Logo" width="200"/>
</p>

# Excel Tool

A desktop app that opens Excel files and displays them as clean, styled HTML tables. Built with Electron.

The goal of this app is to allow users to browse for an Excel file on their machine and view the contents in a nicely formatted table. Files are processed locally — nothing is uploaded.
If a user needs to display a schedule on their desktop constantly as a widget this is a clean and simple way to do so. 

## Features

- Opens `.xlsx` Excel files from your local machine
- Converts Excel sheets into styled tables
- Runs as a standalone macOS desktop application

## Tech Stack

- Electron
- Node.js
- Aspose.Cells for Node.js via C++
- HTML/CSS for display

## Usage

1. Clone the repository:

   ```bash
   git clone https://github.com/YOUR_USERNAME/excel-tool.git
   cd excel-tool
   ```

2. Install dependencies:

   ```bash
   npm install
   ```

3. Run the app:

   ```bash
   npm start
   ```

4. To build the app:
  a. For mac
   ```bash
   npm run dist:mac
   ```
  b. For windows
  ```bash
   npm run dist:win
   ```
  c. For all
   ```bash
   npm run dist:all
   ```

## Notes

- The HTML is generated from JS and then loaded into the app window.
- This app is designed for local file processing only.
