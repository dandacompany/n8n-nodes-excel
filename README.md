# n8n Excel/CSV 파일 I/O 노드

로컬 파일 시스템의 Excel(.xlsx) 및 CSV(.csv) 파일을 데이터베이스처럼 다룰 수 있는 n8n 커스텀 노드입니다. 이 노드를 사용하면 n8n 워크플로우 내에서 직접 파일을 생성, 업로드, 다운로드, 삭제하고, 파일 내의 데이터를 읽고, 추가하고, 수정하는 등의 다양한 작업을 수행할 수 있습니다.

모든 파일은 n8n 인스턴스의 `data` 디렉토리(보통 `.n8n/custom-nodes-data/n8n-nodes-excel/data`)에 안전하게 저장 및 관리됩니다.

## 주요 기능

이 노드는 **파일(File)**과 **데이터(Data)**, 두 가지 리소스를 중심으로 강력한 기능을 제공합니다.

### 파일 (File) 관리

- **생성 (Create):** 원하는 컬럼(헤더)을 지정하여 비어 있는 새로운 Excel 또는 CSV 파일을 생성합니다.
- **업로드 (Upload):** 이전 노드에서 받은 파일(바이너리 데이터)이나 로컬 시스템 경로에 있는 파일을 `data` 디렉토리로 업로드합니다.
- **다운로드 (Download):** `data` 디렉토리에 있는 파일을 워크플로우의 다른 노드로 전달할 수 있도록 바이너리 데이터로 내려받습니다.
- **삭제 (Delete):** `data` 디렉토리에서 특정 파일을 영구적으로 삭제합니다.

### 데이터 (Data) 조작

- **읽기 (Read):** 파일과 시트를 지정하여 내용을 JSON 형식으로 읽어옵니다. 원하는 개수만큼 데이터를 가져올 수 있는 `Limit` 옵션을 지원합니다.
- **행 추가 (Add Row):** 파일과 시트를 선택하면, **해당 시트의 모든 컬럼이 자동으로 입력 필드로 표시됩니다.** 사용자는 각 컬럼에 들어갈 값만 입력하면 간편하게 새로운 행을 추가할 수 있습니다.
- **행 업데이트 (Update Row):** `Key Column`(고유 식별자 컬럼)과 `Key Value`(값)를 기준으로 특정 행을 찾아 데이터를 수정합니다. 행 추가와 마찬가지로, **컬럼 필드가 자동으로 로드되어** 데이터 수정이 매우 편리합니다.
- **데이터 초기화 (Clear Data):** 지정된 시트의 모든 데이터를 삭제하고 헤더(컬럼) 행만 남깁니다.

## 설치 방법

### Docker Compose (권장)

n8n을 Docker Compose로 운영하는 경우, `docker-compose.yml` 파일에 다음과 같이 custom 노드를 위한 볼륨을 추가하고 환경 변수를 설정합니다.

```yaml
version: '3.7'

services:
  n8n:
    # ... 기존 n8n 설정 ...
    environment:
      - NODE_FUNCTION_ALLOW_EXTERNAL=n8n-nodes-excel
      - N8N_CUSTOM_EXTENSIONS=/home/node/.n8n/custom
    volumes:
      - ~/.n8n/custom:/home/node/.n8n/custom
    # ...
```

yml 파일 수정 후, 다음 명령어로 n8n을 재시작합니다.

```bash
docker-compose up -d --force-recreate
```

n8n 컨테이너에 접속하여 노드를 설치합니다.

```bash
docker exec -it <n8n-container-name> /bin/sh
# 컨테이너 내부에서 실행
npm install n8n-nodes-excel
```

마지막으로 n8n 컨테이너를 재시작하면 노드가 적용됩니다.

### 로컬 환경 (npm)

로컬 머신에 n8n을 직접 설치한 경우, 다음 명령어로 간단하게 설치할 수 있습니다.

1. n8n의 커스텀 노드 디렉토리로 이동합니다.

    ```bash
    cd ~/.n8n/nodes
    ```

2. npm을 사용하여 노드를 설치합니다.

    ```bash
    npm install n8n-nodes-excel
    ```

3. n8n을 재시작합니다.

설치가 완료되면 n8n 편집기의 노드 패널에서 `Excel/CSV File IO` 노드를 검색하여 사용할 수 있습니다.

## 개발자용 설정

이 노드의 소스 코드를 직접 수정하고 개발하려는 경우 다음 안내를 따르세요.

### 필요 사항

- [Node.js](https://nodejs.org/) (v16 이상 권장)
- [n8n](https://n8n.io/) (로컬 개발용으로 설치)

### 빌드 및 테스트

1. 이 저장소를 클론하고 의존성을 설치합니다.

    ```bash
    git clone https://github.com/your-repo/n8n-nodes-excel.git
    cd n8n-nodes-excel
    npm install
    ```

2. 소스 코드를 빌드합니다.

    ```bash
    npm run build
    ```

3. 로컬 n8n에 개발 중인 노드를 연결합니다.

    ```bash
    # n8n 프로젝트의 루트 디렉토리에서 실행
    npm link /path/to/your/n8n-nodes-excel
    ```

4. 개발 모드로 n8n을 시작합니다.

    ```bash
    n8n start --dev
    ```
