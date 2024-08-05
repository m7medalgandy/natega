import { read, utils } from 'xlsx';

const R2_BUCKET_NAME = 'ntega';
const EXCEL_FILE_NAME = 'students-results.xlsx';

let studentsData = null;

addEventListener('fetch', event => {
  event.respondWith(handleRequest(event.request));
});

async function handleRequest(request) {
  if (request.method === 'POST') {
    return await handleSearch(request);
  }
  return await serveHTML();
}

async function readExcelFromR2() {
  const bucket = R2_BUCKET.get(R2_BUCKET_NAME);
  const object = await bucket.get(EXCEL_FILE_NAME);
  if (object === null) {
    throw new Error('الملف غير موجود');
  }
  const arrayBuffer = await object.arrayBuffer();
  const workbook = read(arrayBuffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  return utils.sheet_to_json(sheet);
}

async function initializeData() {
  if (studentsData === null) {
    const excelData = await readExcelFromR2();
    studentsData = excelData.reduce((acc, row) => {
      acc[row.seatNumber] = row;
      return acc;
    }, {});
  }
}

async function handleSearch(request) {
  await initializeData();
  const formData = await request.formData();
  const seatNumber = formData.get('seatNumber');
  
  if (seatNumber in studentsData) {
    const student = studentsData[seatNumber];
    return new Response(JSON.stringify(student), {
      headers: { 'Content-Type': 'application/json' }
    });
  } else {
    return new Response(JSON.stringify({ error: 'لم يتم العثور على الطالب' }), {
      status: 404,
      headers: { 'Content-Type': 'application/json' }
    });
  }
}

async function serveHTML() {
  const html = `
<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>البحث برقم الجلوس</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; }
        form { margin-bottom: 20px; }
        input, button { font-size: 16px; padding: 5px; }
    </style>
</head>
<body>
    <h1>البحث برقم الجلوس</h1>
    <form id="searchForm">
        <input type="text" id="seatNumber" name="seatNumber" placeholder="أدخل رقم الجلوس" required>
        <button type="submit">بحث</button>
    </form>
    <div id="result"></div>

    <script>
        document.getElementById('searchForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            const formData = new FormData(e.target);
            const response = await fetch('/', {
                method: 'POST',
                body: formData
            });
            const result = await response.json();
            const resultDiv = document.getElementById('result');
            if (response.ok) {
                resultDiv.innerHTML = `
                    <h2>نتيجة البحث:</h2>
                    <p>رقم الجلوس: ${result.seatNumber}</p>
                    <p>الاسم: ${result.name}</p>
                    <p>الدرجة: ${result.grade}</p>
                `;
            } else {
                resultDiv.innerHTML = `<p>خطأ: ${result.error}</p>`;
            }
        });
    </script>
</body>
</html>
  `;
  
  return new Response(html, {
    headers: { 'Content-Type': 'text/html; charset=utf-8' }
  });
}
