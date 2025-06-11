const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const cors = require('cors');
const multer = require('multer');

const app = express();
const port = 3000;

// Excel 파일을 저장하고 관리할 기본 디렉토리
const dataDir = path.join(__dirname, 'data');

// 서버 시작 시 data 디렉토리가 없으면 생성
if (!fs.existsSync(dataDir)) {
    fs.mkdirSync(dataDir, { recursive: true });
}

// Multer 설정: data 디렉토리에 원본 파일 이름으로 저장
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, dataDir)
    },
    filename: function (req, file, cb) {
        // 파일 이름이 중복될 경우 덮어쓰지 않도록 처리합니다.
        // 여기서는 간단하게 원본 이름을 사용하지만, 실제 프로덕션에서는
        // 중복을 피하기 위해 파일 이름에 타임스탬프 등을 추가하는 것이 좋습니다.
        cb(null, file.originalname)
    }
});

const fileFilter = (req, file, cb) => {
    if (file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || file.mimetype === 'text/csv') {
        cb(null, true);
    } else {
        cb(new Error('Only .xlsx and .csv files are allowed!'), false);
    }
};

const upload = multer({ storage: storage, fileFilter: fileFilter });

app.use(cors());
app.use(express.json());

app.get('/files', (req, res) => {
    try {
        const files = fs.readdirSync(dataDir).filter(file => file.endsWith('.xlsx') || file.endsWith('.csv'));
        res.json(files);
    } catch (error) {
        console.error(error);
        res.status(500).send('파일 목록을 가져오는 중 오류가 발생했습니다.');
    }
});

/**
 * @api {post} /upload 로컬 파일을 서버에 업로드
 * @apiDescription 로컬 컴퓨터의 Excel 파일을 서버의 data 디렉토리에 업로드합니다.
 * @apiName UploadExcel
 * @apiGroup LocalExcelDB
 */
app.post('/upload', upload.single('excel'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('업로드할 파일이 없습니다.');
    }
    res.json({ message: `파일이 성공적으로 업로드되었습니다: ${req.file.filename}` });
});

/**
 * @api {get} /download/:fileName 서버의 파일을 다운로드
 * @apiDescription 서버 data 디렉토리의 지정된 파일을 다운로드합니다.
 * @apiName DownloadExcel
 * @apiGroup LocalExcelDB
 */
app.get('/download/:fileName', (req, res) => {
    const { fileName } = req.params;
    const filePath = path.join(dataDir, fileName);

    if (fs.existsSync(filePath)) {
        res.download(filePath, fileName, (err) => {
            if (err) {
                console.error("File download error:", err);
                res.status(500).send('파일 다운로드 중 오류가 발생했습니다.');
            }
        });
    } else {
        res.status(404).send('파일을 찾을 수 없습니다.');
    }
});

/**
 * @api {post} /delete 서버의 파일 삭제
 * @apiDescription 서버 data 디렉토리의 지정된 파일을 삭제합니다.
 * @apiName DeleteExcel
 * @apiGroup LocalExcelDB
 */
app.post('/delete', (req, res) => {
    const { fileName } = req.body;
    if (!fileName) {
        return res.status(400).send('삭제할 파일 이름을 제공해야 합니다.');
    }
    const filePath = path.join(dataDir, fileName);

    if (fs.existsSync(filePath)) {
        try {
            fs.unlinkSync(filePath);
            res.json({ message: `파일이 성공적으로 삭제되었습니다: ${fileName}` });
        } catch (error) {
            console.error(error);
            res.status(500).send('파일 삭제 중 오류가 발생했습니다.');
        }
    } else {
        res.status(404).send('삭제할 파일을 찾을 수 없습니다.');
    }
});

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
 * @apiParam {String} body.limit 반환할 행 수. 기본값은 5.
 *
 * @apiSuccess {String} fileName 파일 이름.
 * @apiSuccess {Object[]} data Excel 파일의 첫번째 시트 데이터.
 */
app.post('/read', (req, res) => {
    const { fileName, sheetName: requestedSheetName, limit = 5 } = req.body;
    if (!fileName) {
        return res.status(400).send('fileName을 제공해야 합니다.');
    }

    const filePath = path.join(dataDir, fileName);

    try {
        if (!fs.existsSync(filePath)) {
            return res.status(404).send('파일을 찾을 수 없습니다.');
        }

        const workbook = xlsx.readFile(filePath);
        const sheetName = requestedSheetName || workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        if (!sheet) {
            return res.status(404).send(`'${sheetName}' 시트를 찾을 수 없습니다.`);
        }

        // 1. 빈 셀을 빈 문자열로 읽도록 defval 옵션 추가
        const jsonData = xlsx.utils.sheet_to_json(sheet, { defval: "" });

        if (jsonData.length === 0) {
            return res.json([]);
        }

        // 2. 데이터 클리닝 로직
        const cleanedData = jsonData.filter(row => {
            const values = Object.values(row);
            // 모든 값이 비어있거나, 키가 '__EMPTY'로 시작하는 자동 생성된 키만 있는 경우 제외
            return values.some(v => v !== "") && Object.keys(row).some(k => !k.startsWith('__EMPTY'));
        }).map(row => {
            // '__EMPTY' 키를 가진 속성 제거
            const newRow = {};
            for (const key in row) {
                if (!key.startsWith('__EMPTY')) {
                    newRow[key] = row[key];
                }
            }
            return newRow;
        });

        const finalData = cleanedData.slice(0, Number(limit));

        res.status(200).json(finalData);
    } catch (error) {
        console.error('Error reading file:', error);
        res.status(500).send('파일을 읽는 중 오류가 발생했습니다.');
    }
});

/**
 * @api {post} /create 로컬에 Excel 파일 생성
 * @apiDescription CSV 형식의 데이터 문자열을 받아 새로운 Excel 파일을 생성하고 서버의 data 디렉토리에 저장합니다.
 * @apiName CreateExcel
 * @apiGroup LocalExcelDB
 *
 * @apiParam {Object} body Request body.
 * @apiParam {String} body.fileName 생성할 파일 이름.
 * @apiParam {String} body.csvData CSV 형식의 데이터 문자열. 첫 줄은 헤더로 사용됩니다.
 *
 * @apiSuccess {String} message 성공 메시지.
 */
app.post('/create', (req, res) => {
    try {
        const { fileName, csvData } = req.body;

        if (!fileName || !csvData) {
            return res.status(400).send('fileName과 csvData를 제공해야 합니다.');
        }

        const filePath = path.join(dataDir, fileName);
        if (fs.existsSync(filePath)) {
            return res.status(409).send('이미 해당 이름의 파일이 존재합니다.');
        }

        // CSV 데이터를 파싱하여 2차원 배열로 변환합니다.
        // 참고: 이 파서는 간단한 CSV만 지원하며, 쉼표를 포함한 셀은 처리하지 못할 수 있습니다.
        const rows = csvData.trim().split('\n');
        const aoa = rows.map(row => row.split(','));

        const workbook = xlsx.utils.book_new();
        const sheet = xlsx.utils.aoa_to_sheet(aoa);
        xlsx.utils.book_append_sheet(workbook, sheet, 'Sheet1');

        xlsx.writeFile(workbook, filePath);

        res.status(201).json({ message: `파일이 성공적으로 생성되었습니다: ${fileName}` });

    } catch (error) {
        console.error(error);
        res.status(500).send('파일 생성 중 오류가 발생했습니다.');
    }
});

/**
 * @api {post} /sheets 파일의 시트 목록 가져오기
 * @apiDescription 지정된 Excel 파일의 모든 시트 이름 목록을 반환합니다.
 */
app.post('/sheets', (req, res) => {
    const { fileName } = req.body;
    if (!fileName) return res.status(400).send('fileName을 제공해야 합니다.');
    const filePath = path.join(dataDir, fileName);
    if (!fs.existsSync(filePath)) return res.status(404).send('파일을 찾을 수 없습니다.');
    try {
        const workbook = xlsx.readFile(filePath);
        res.json(workbook.SheetNames);
    } catch (error) {
        res.status(500).send('시트 목록을 가져오는 중 오류 발생');
    }
});

/**
 * @api {post} /headers 시트의 헤더 목록 가져오기
 * @apiDescription 지정된 파일과 시트의 첫 번째 행(헤더)을 배열로 반환합니다.
 */
app.post('/headers', (req, res) => {
    const { fileName, sheetName } = req.body;
    if (!fileName || !sheetName) return res.status(400).send('fileName과 sheetName을 제공해야 합니다.');
    const filePath = path.join(dataDir, fileName);
    if (!fs.existsSync(filePath)) return res.status(404).send('파일을 찾을 수 없습니다.');
    try {
        const workbook = xlsx.readFile(filePath);
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) {
            return res.status(404).send('시트를 찾을 수 없습니다.');
        }

        const headers = xlsx.utils.sheet_to_json(sheet, { header: 1 })[0] || [];
        const cleanedHeaders = headers.filter(h => h && !String(h).startsWith('__EMPTY'));

        res.status(200).json(cleanedHeaders);
    } catch (error) {
        console.error('Error reading headers:', error);
        res.status(500).send('헤더를 읽는 중 오류가 발생했습니다.');
    }
});

/**
 * @api {post} /update-row Key를 기준으로 행 수정하기
 * @apiDescription 지정된 Key와 값을 기준으로 행을 찾아 데이터를 수정합니다.
 */
app.post('/update-row', (req, res) => {
    const { fileName, sheetName, keyColumn, keyValue, rowData } = req.body;
    if (!fileName || !sheetName || !keyColumn || keyValue === undefined || !rowData) {
        return res.status(400).send('필수 파라미터가 누락되었습니다.');
    }

    const filePath = path.join(dataDir, fileName);
    if (!fs.existsSync(filePath)) return res.status(404).send('파일을 찾을 수 없습니다.');

    try {
        const workbook = xlsx.readFile(filePath);
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) return res.status(404).send('시트를 찾을 수 없습니다.');

        const jsonData = xlsx.utils.sheet_to_json(sheet, { defval: "" });

        // Find the row index to update
        const rowIndex = jsonData.findIndex(row => String(row[keyColumn]) === String(keyValue));

        if (rowIndex === -1) {
            return res.status(404).send(`'${keyColumn}' 열에서 '${keyValue}' 값을 가진 행을 찾을 수 없습니다.`);
        }

        // Update the row with new data
        jsonData[rowIndex] = { ...jsonData[rowIndex], ...rowData };

        // Create a new sheet with the updated data
        const newSheet = xlsx.utils.json_to_sheet(jsonData);

        // Replace the old sheet with the new one
        workbook.Sheets[sheetName] = newSheet;

        xlsx.writeFile(workbook, filePath);

        res.json({ message: `행이 성공적으로 수정되었습니다.` });
    } catch (error) {
        console.error(error);
        res.status(500).send('행 수정 중 오류 발생');
    }
});

/**
 * @api {post} /add-row 시트에 행 추가하기
 * @apiDescription 지정된 파일과 시트의 마지막에 새로운 데이터 행을 추가합니다.
 */
app.post('/add-row', (req, res) => {
    const { fileName, sheetName, rowData } = req.body;
    if (!fileName || !sheetName || !rowData) return res.status(400).send('필수 파라미터가 누락되었습니다.');

    const filePath = path.join(dataDir, fileName);
    if (!fs.existsSync(filePath)) return res.status(404).send('파일을 찾을 수 없습니다.');

    try {
        const workbook = xlsx.readFile(filePath);
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) return res.status(404).send('시트를 찾을 수 없습니다.');

        // origin: -1 은 시트의 마지막에 데이터를 추가하라는 의미입니다.
        // sheet_add_json 은 헤더를 기반으로 JSON 객체를 추가하므로 rowData가 {header: value} 형태여야 합니다.
        xlsx.utils.sheet_add_json(sheet, [rowData], { skipHeader: true, origin: -1 });
        xlsx.writeFile(workbook, filePath);

        res.json({ message: `'${sheetName}' 시트에 행이 성공적으로 추가되었습니다.` });
    } catch (error) {
        console.error(error);
        res.status(500).send('행 추가 중 오류 발생');
    }
});

/**
 * @api {post} /clear-data Clear all data from a sheet, leaving only the header row.
 * @apiName ClearData
 * @apiGroup Excel
 *
 * @apiParam {Object} body Request body.
 * @apiParam {String} body.fileName The name of the Excel file in the data directory.
 * @apiParam {String} body.sheetName The name of the sheet to clear.
 *
 * @apiSuccess {String} message Success message.
 */
app.post('/clear-data', (req, res) => {
    const { fileName, sheetName } = req.body;
    if (!fileName || !sheetName) {
        return res.status(400).send('fileName과 sheetName을 모두 제공해야 합니다.');
    }

    const filePath = path.join(dataDir, fileName);

    try {
        if (!fs.existsSync(filePath)) {
            return res.status(404).send('파일을 찾을 수 없습니다.');
        }

        const workbook = xlsx.readFile(filePath);
        const sheet = workbook.Sheets[sheetName];

        if (!sheet) {
            return res.status(404).send(`'${sheetName}' 시트를 찾을 수 없습니다.`);
        }

        // Get headers
        const headers = (xlsx.utils.sheet_to_json(sheet, { header: 1 })[0] || []).filter(h => h && !String(h).startsWith('__EMPTY'));

        // Create a new sheet with only headers
        const newSheet = xlsx.utils.aoa_to_sheet([headers]);

        // Replace the old sheet with the new one
        workbook.Sheets[sheetName] = newSheet;

        // Write the updated workbook back to the file
        xlsx.writeFile(workbook, filePath);

        res.status(200).send({ message: `'${sheetName}' 시트의 모든 데이터가 삭제되었습니다.` });
    } catch (error) {
        console.error('Error clearing data:', error);
        res.status(500).send('데이터를 삭제하는 중 오류가 발생했습니다.');
    }
});

app.listen(port, () => {
    console.log(`서버가 http://localhost:${port} 에서 실행중입니다.`);
    console.log(`Excel 파일은 '${dataDir}' 디렉토리에서 관리됩니다.`);
}); 