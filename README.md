1. import js

```
  <script src="xlsx.full.min.js"></script>
  <script src="excletotxt.js"></script>
```

2. add script

```
  <div id="download-link1"></div>
  <div id="download-link2"></div>
  <script>
	const filePath1 = '/test2.xlsx';
	const sheetName1 = 'Sheet2';
	const startCell1 = 'A2';
	const fileName1 = '12345.txt';
	const linkText1 = 'Download File 1';
	const targetDivId1 = 'download-link1';
	handleExcelFile(filePath1, sheetName1, startCell1, fileName1, targetDivId1, linkText1);

	const filePath2 = '/test2.xlsx';
	const sheetName2 = 'Sheet1';
	const startCell2 = 'A2';
	const fileName2 = '234.txt';
	const linkText2 = 'Download File 2';
	const targetDivId2 = 'download-link2';
	handleExcelFile(filePath2, sheetName2, startCell2, fileName2, targetDivId2, linkText2);
	
......more and more
  </script>
```
