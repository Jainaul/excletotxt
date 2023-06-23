// Read the contents of a specified range of cells from an Excel file
function readExcelCellRange(filePath, sheetName, startCell) {
  return new Promise(resolve => {
    const oReq = new XMLHttpRequest();
    oReq.open('GET', filePath, true);
    oReq.responseType = 'arraybuffer';

    oReq.onload = function (e) {
      const arraybuffer = oReq.response;
      const data = new Uint8Array(arraybuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const worksheet = workbook.Sheets[sheetName];

      const lastCell = worksheet['!ref'].split(':')[1];
      const lastRow = XLSX.utils.decode_cell(lastCell).r;

      const nextRow = lastRow + 1;

      const range = XLSX.utils.decode_range(`${startCell}:A${nextRow}`);
      const cellValues = [];

      for (let row = range.s.r; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          const cellValue = worksheet[cellAddress]?.v;
          cellValues.push(cellValue);
        }
      }

      resolve(cellValues);
    };

    oReq.send();
  });
}

// Generate download links and create download files
function generateDownloadLink(content, fileName, targetDivId, linkText) {
  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);

  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  link.innerHTML = linkText;

  document.getElementById(targetDivId).appendChild(link);
}

// Process Excel files and generate download links
function handleExcelFile(filePath, sheetName, startCell, fileName, targetDivId, linkText) {
  readExcelCellRange(filePath, sheetName, startCell)
    .then(content => generateDownloadLink(content.join('\n'), fileName, targetDivId, linkText));
}
