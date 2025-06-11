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
    res.send('Excel IO API 서버가 실행중입니다.');
});

/**
 * @api {post} /upload Excel 파일 업로드 및 JSON 변환
 * @apiDescription 첫 행을 헤더로 사용하여 Excel 데이터를 JSON 레코드 배열로 변환합니다.
 * @apiName UploadExcel
 * @apiGroup Excel
 *
 * @apiParam {File} excel form-data의 'excel' 키로 전송되는 파일.
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
 * @apiDescription JSON 데이터를 Excel 파일로 생성합니다. 헤더 순서를 지정할 수 있습니다.
 * @apiName DownloadExcel
 * @apiGroup Excel
 *
 * @apiParam {Object} body Request body.
 * @apiParam {Object[]} body.data 변환할 JSON 객체 배열.
 * @apiParam {String[]} [body.headers] (선택) Excel 파일의 헤더로 사용할 문자열 배열. 지정하면 해당 순서대로 열이 생성됩니다.
 * @apiParam {String} [fileName='output.xlsx'] (쿼리 파라미터) 다운로드 될 파일 이름.
 *
 * @apiSuccess {File} excel 생성된 Excel 파일.
 */
app.post('/download', (req, res) => {
    try {
        const { data, headers } = req.body;

        if (!data || !Array.isArray(data) || data.length === 0) {
            return res.status(400).send('Excel 파일로 변환할 유효한 JSON 데이터가 필요합니다.');
        }

        const workbook = xlsx.utils.book_new();
        let sheet;

        if (headers && Array.isArray(headers) && headers.length > 0) {
            // headers 옵션을 사용하여 헤더 순서를 지정합니다.
            sheet = xlsx.utils.json_to_sheet(data, { header: headers });
        } else {
            // 헤더가 없으면 기존 방식대로 객체의 키를 헤더로 사용합니다.
            sheet = xlsx.utils.json_to_sheet(data);
        }

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

/**
 * @api {post} /modify 기존 Excel 파일 수정
 * @apiDescription 기존 Excel 파일의 특정 시트, 특정 셀 범위에 데이터를 쓰거나 수정합니다.
 * @apiName ModifyExcel
 * @apiGroup Excel
 *
 * @apiParam {File} excel form-data의 'excel' 키로 전송되는 수정할 파일.
 * @apiParam {String} data JSON.stringify 된 2차원 배열 데이터 (e.g., '[["Col1", "Col2"], [1, 2]]').
 * @apiParam {String} [sheet] (선택) 수정할 시트 이름. 지정하지 않으면 첫 번째 시트. 시트가 없으면 새로 생성됩니다.
 * @apiParam {String} [origin="A1"] (선택) 데이터 쓰기를 시작할 셀 주소 (e.g., "A1", "C5").
 *
 * @apiSuccess {File} excel 수정된 Excel 파일.
 */
app.post('/modify', upload.single('excel'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('수정할 파일이 업로드되지 않았습니다.');
    }
    if (!req.body.data) {
        return res.status(400).send('쓸 데이터가 없습니다.');
    }

    try {
        const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
        let dataToWrite;
        try {
            dataToWrite = JSON.parse(req.body.data);
            if (!Array.isArray(dataToWrite)) throw new Error();
        } catch (e) {
            return res.status(400).send('data는 유효한 JSON 배열 문자열이어야 합니다.');
        }

        const sheetName = req.body.sheet || workbook.SheetNames[0];
        let sheet = workbook.Sheets[sheetName];

        // 시트가 없으면 새로 생성
        if (!sheet) {
            sheet = xlsx.utils.aoa_to_sheet([]);
            xlsx.utils.book_append_sheet(workbook, sheet, sheetName);
        }

        xlsx.utils.sheet_add_aoa(sheet, dataToWrite, { origin: req.body.origin || 'A1' });

        const fileName = req.query.fileName || 'modified.xlsx';

        res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
        res.send(buffer);

    } catch (error) {
        console.error(error);
        res.status(500).send('파일 수정 중 오류가 발생했습니다.');
    }
});


app.listen(port, () => {
    console.log(`서버가 http://localhost:${port} 에서 실행중입니다.`);
}); 