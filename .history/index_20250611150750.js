const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const port = 3000;

// Excel 파일을 저장하고 관리할 기본 디렉토리
const dataDir = path.join(__dirname, 'data');

// 서버 시작 시 data 디렉토리가 없으면 생성
if (!fs.existsSync(dataDir)) {
    fs.mkdirSync(dataDir, { recursive: true });
}

app.use(express.json());

// 기본 라우트
app.get('/', (req, res) => {
    res.send('로컬 Excel 파일 DB API 서버가 실행중입니다.');
});

/**
 * @api {post} /read 로컬 Excel 파일 읽기
 * @apiDescription 지정된 Excel 파일을 읽어 JSON 레코드 배열로 반환합니다.
 * @apiName ReadExcel
 * @apiGroup LocalExcelDB
 *
 * @apiParam {Object} body Request body.
 * @apiParam {String} body.fileName data 디렉토리 내의 Excel 파일 이름.
 *
 * @apiSuccess {Object[]} data Excel 파일의 첫번째 시트 데이터.
 */
app.post('/read', (req, res) => {
    const { fileName } = req.body;
    if (!fileName) {
        return res.status(400).send('fileName을 제공해야 합니다.');
    }

    const filePath = path.join(dataDir, fileName);

    try {
        if (!fs.existsSync(filePath)) {
            return res.status(404).send('파일을 찾을 수 없습니다.');
        }

        const workbook = xlsx.readFile(filePath);
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
 * @api {post} /create 로컬에 Excel 파일 생성
 * @apiDescription JSON 데이터를 받아 새로운 Excel 파일을 생성하고 서버의 data 디렉토리에 저장합니다.
 * @apiName CreateExcel
 * @apiGroup LocalExcelDB
 *
 * @apiParam {Object} body Request body.
 * @apiParam {String} body.fileName 생성할 파일 이름.
 * @apiParam {Object[]} body.data Excel 파일에 쓸 데이터.
 * @apiParam {String[]} [body.headers] (선택) 헤더로 사용할 문자열 배열.
 *
 * @apiSuccess {String} message 성공 메시지.
 */
app.post('/create', (req, res) => {
    try {
        const { fileName, data, headers } = req.body;

        if (!fileName || !data || !Array.isArray(data)) {
            return res.status(400).send('fileName과 data 배열을 제공해야 합니다.');
        }

        const filePath = path.join(dataDir, fileName);
        if (fs.existsSync(filePath)) {
            return res.status(409).send('이미 해당 이름의 파일이 존재합니다. 수정을 원하시면 /modify를 사용하세요.');
        }

        const workbook = xlsx.utils.book_new();
        const sheet = xlsx.utils.json_to_sheet(data, { header: headers });
        xlsx.utils.book_append_sheet(workbook, sheet, 'Sheet1');

        xlsx.writeFile(workbook, filePath);

        res.status(201).json({ message: `파일이 성공적으로 생성되었습니다: ${fileName}` });

    } catch (error) {
        console.error(error);
        res.status(500).send('파일 생성 중 오류가 발생했습니다.');
    }
});

/**
 * @api {post} /modify 로컬 Excel 파일 수정
 * @apiDescription 기존 Excel 파일의 특정 위치에 데이터를 쓰거나 수정합니다.
 * @apiName ModifyExcel
 * @apiGroup LocalExcelDB
 *
 * @apiParam {Object} body Request body.
 * @apiParam {String} body.fileName 수정할 파일 이름.
 * @apiParam {Array[]} body.data 파일에 쓸 2차원 배열 데이터.
 * @apiParam {String} [body.sheet] (선택) 수정할 시트 이름. 없으면 첫 번째 시트.
 * @apiParam {String} [body.origin="A1"] (선택) 데이터 쓰기를 시작할 셀 주소.
 *
 * @apiSuccess {String} message 성공 메시지.
 */
app.post('/modify', (req, res) => {
    const { fileName, data, sheet: sheetNameParam, origin } = req.body;

    if (!fileName || !data || !Array.isArray(data)) {
        return res.status(400).send('fileName과 data 배열을 제공해야 합니다.');
    }

    const filePath = path.join(dataDir, fileName);

    try {
        if (!fs.existsSync(filePath)) {
            return res.status(404).send('수정할 파일을 찾을 수 없습니다. 먼저 /create로 생성하세요.');
        }

        const workbook = xlsx.readFile(filePath);
        const sheetName = sheetNameParam || workbook.SheetNames[0];
        let sheet = workbook.Sheets[sheetName];

        if (!sheet) {
            sheet = xlsx.utils.aoa_to_sheet([]);
            xlsx.utils.book_append_sheet(workbook, sheet, sheetName);
        }

        xlsx.utils.sheet_add_aoa(sheet, data, { origin: origin || 'A1' });
        xlsx.writeFile(workbook, filePath);

        res.json({ message: `파일이 성공적으로 수정되었습니다: ${fileName}` });

    } catch (error) {
        console.error(error);
        res.status(500).send('파일 수정 중 오류가 발생했습니다.');
    }
});


app.listen(port, () => {
    console.log(`서버가 http://localhost:${port} 에서 실행중입니다.`);
    console.log(`Excel 파일은 '${dataDir}' 디렉토리에서 관리됩니다.`);
}); 