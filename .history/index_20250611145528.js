const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');

const app = express();
const port = 3000;

// Multer 설정을 통해 메모리에 파일을 저장합니다.
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// 기본 라우트
app.get('/', (req, res) => {
    res.send('Excel IO API 서버가 실행중입니다. /upload 로 POST 요청을 보내 파일을 업로드하거나, /download 로 POST 요청을 보내 파일을 다운로드하세요.');
});

/**
 * @api {post} /upload Excel 파일 업로드 및 JSON 변환
 * @apiName UploadExcel
 * @apiGroup Excel
 *
 * @apiParam {File} excel 파일. form-data 에서 'excel' 이라는 키로 전송해야 합니다.
 *
 * @apiSuccess {Object[]} data Excel 파일의 첫번째 시트 데이터.
 */
app.post('/upload', upload.single('excel'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('파일이 업로드되지 않았습니다.');
    }

    try {
        const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(sheet);

        res.json(data);
    } catch (error) {
        console.error(error);
        res.status(500).send('파일 처리 중 오류가 발생했습니다.');
    }
});

/**
 * @api {post} /download JSON 데이터를 Excel 파일로 변환 및 다운로드
 * @apiName DownloadExcel
 * @apiGroup Excel
 *
 * @apiParam {Object[]} data JSON 배열 데이터. request body 에 담아 전송해야 합니다.
 * @apiParam {String} [fileName='output.xlsx'] (쿼리 파라미터) 다운로드 될 파일 이름.
 *
 * @apiSuccess {File} excel 생성된 Excel 파일.
 */
app.post('/download', (req, res) => {
    try {
        const data = req.body;
        if (!data || !Array.isArray(data) || data.length === 0) {
            return res.status(400).send('Excel 파일로 변환할 유효한 JSON 데이터가 필요합니다.');
        }

        const sheet = xlsx.utils.json_to_sheet(data);
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, sheet, 'Sheet1');

        const fileName = req.query.fileName || 'output.xlsx';

        // 브라우저에서 다운로드할 수 있도록 헤더 설정
        res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
        res.send(buffer);

    } catch (error) {
        console.error(error);
        res.status(500).send('파일 생성 중 오류가 발생했습니다.');
    }
});


app.listen(port, () => {
    console.log(`서버가 http://localhost:${port} 에서 실행중입니다.`);
}); 