document.getElementById('openBtn').addEventListener('click', async () => {
  const result = await window.api.openFile();

  if (!result.canceled) {
    document.getElementById('fileContent').innerText = result.content;
    console.log('File Content:', result.content);
  }
});