# 로컬 Excel/CSV 파일 데이터베이스 API 서버

Node.js와 Express를 사용하여 로컬 파일 시스템의 Excel(.xlsx) 및 CSV(.csv) 파일을 데이터베이스처럼 다루는 API 서버입니다. 서버 내의 `data` 디렉토리에 있는 파일을 읽고, 생성하고, 수정할 수 있는 포괄적인 엔드포인트를 제공합니다.

## 시작하기

### 필요 사항

- [Node.js](https://nodejs.org/) (v14 이상 권장)

### 설치 및 실행

1. 저장소를 클론하거나 코드를 다운로드합니다.
2. 프로젝트 루트 디렉토리에서 의존성을 설치합니다.

    ```bash
    npm install
    ```

3. 서버를 시작합니다.

    ```bash
    node index.js
    ```

4. 서버는 `http://localhost:3000` 에서 실행됩니다.
5. 모든 파일은 프로젝트 루트에 자동으로 생성되는 `data/` 디렉토리 내에서 관리됩니다.

## API 엔드포인트

### 파일 관리

#### 1. 파일 목록 조회

`data` 디렉토리에 있는 모든 Excel(.xlsx) 및 CSV(.csv) 파일의 목록을 반환합니다.

- **URL**: `/files`
- **Method**: `GET`
- **성공 응답**: `["products.xlsx", "sales_data.csv"]`
- **`curl` 예시**: `curl http://localhost:3000/files`

#### 2. 파일 업로드

새로운 Excel 또는 CSV 파일을 `data` 디렉토리에 업로드합니다.

- **URL**: `/upload`
- **Method**: `POST`
- **Request Body**: `multipart/form-data`
  - `excel`: 업로드할 파일
- **`curl` 예시**:

    ```bash
    curl -X POST -F "excel=@/path/to/your/file.xlsx" http://localhost:3000/upload
    ```

#### 3. 파일 다운로드

지정된 파일을 다운로드합니다.

- **URL**: `/download/:fileName`
- **Method**: `GET`
- **`curl` 예시**:

    ```bash
    curl -o downloaded_file.xlsx http://localhost:3000/download/products.xlsx
    ```

#### 4. 새 파일 생성

CSV 형식의 텍스트 데이터를 받아 새 Excel 또는 CSV 파일을 생성합니다.

- **URL**: `/create`
- **Method**: `POST`
- **Request Body**:

    ```json
    {
      "fileName": "new_products.xlsx",
      "csvData": "ID,Name,Price\n1,Keyboard,75\n2,Monitor,200"
    }
    ```

- **`curl` 예시**:

    ```bash
    curl -X POST -H "Content-Type: application/json" \
    -d '{"fileName": "products.csv", "csvData": "ID,Name\n1,Apple\n2,Banana"}' \
    http://localhost:3000/create
    ```

#### 5. 파일 삭제

지정된 파일을 삭제합니다.

- **URL**: `/delete`
- **Method**: `POST`
- **Request Body**: `{"fileName": "products.xlsx"}`
- **`curl` 예시**:

    ```bash
    curl -X POST -H "Content-Type: application/json" \
    -d '{"fileName": "products.xlsx"}' \
    http://localhost:3000/delete
    ```

### 데이터 및 스키마 조회

#### 1. 시트 목록 조회

Excel 파일 내의 모든 시트 이름을 배열로 반환합니다. (CSV 파일의 경우 기본 시트 이름 반환)

- **URL**: `/sheets`
- **Method**: `POST`
- **Request Body**: `{"fileName": "products.xlsx"}`

#### 2. 헤더 조회

지정된 시트의 헤더(첫 번째 행)를 배열로 반환합니다.

- **URL**: `/headers`
- **Method**: `POST`
- **Request Body**: `{"fileName": "products.xlsx", "sheetName": "Sheet1"}`

#### 3. 데이터 읽기

파일의 데이터를 JSON 배열로 반환합니다. 행 수를 제한할 수 있습니다.

- **URL**: `/read`
- **Method**: `POST`
- **Request Body**:

    ```json
    {
      "fileName": "products.xlsx",
      "sheetName": "Sheet1",
      "limit": 10
    }
    ```

  - `sheetName` (String, Optional): `.xlsx` 파일에 필요.
  - `limit` (Number, Optional): 반환할 행의 수. 기본값: 5.

### 데이터 조작

#### 1. 행 추가

지정된 시트에 새로운 데이터 행을 추가합니다.

- **URL**: `/add-row`
- **Method**: `POST`
- **Request Body**:

    ```json
    {
      "fileName": "products.xlsx",
      "sheetName": "Sheet1",
      "rowData": { "ID": 3, "Name": "Desk", "Price": 150 }
    }
    ```

#### 2. 행 업데이트 (키 기준)

Key-Value 쌍을 기준으로 특정 행의 데이터를 찾아 업데이트합니다.

- **URL**: `/update-row`
- **Method**: `POST`
- **Request Body**:

    ```json
    {
      "fileName": "products.xlsx",
      "sheetName": "Sheet1",
      "keyColumn": "ID",
      "keyValue": "2",
      "rowData": { "ID": 2, "Name": "Monitor HD", "Price": 250 }
    }
    ```

#### 3. 데이터 초기화

지정된 시트의 모든 데이터를 삭제하고 헤더 행만 남깁니다.

- **URL**: `/clear-data`
- **Method**: `POST`
- **Request Body**: `{"fileName": "products.xlsx", "sheetName": "Sheet1"}`
- **`curl` 예시**:

    ```bash
    curl -X POST -H "Content-Type: application/json" \
    -d '{"fileName": "products.xlsx", "sheetName": "Sheet1"}' \
    http://localhost:3000/clear-data
    ```
