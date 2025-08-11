# n8n-nodes-excel 프로젝트 규칙

## 프로젝트 개요
Excel 및 CSV 파일을 데이터베이스처럼 다룰 수 있는 n8n 커스텀 노드 프로젝트입니다.

## 기술 스택
- **Runtime**: Node.js (ES2019)
- **Language**: TypeScript 4.8.4
- **Build Tool**: TypeScript Compiler + Gulp (아이콘 복사)
- **Test Framework**: Jest with ts-jest
- **Linting**: ESLint with n8n-nodes-base plugin
- **Core Dependencies**:
  - n8n-workflow: n8n 노드 인터페이스
  - xlsx: Excel 파일 처리
  - fast-csv: CSV 파일 처리

## 프로젝트 구조
```
n8n-nodes-excel/
├── nodes/              # 소스 코드 루트 (TypeScript)
│   └── Excel/
│       ├── Excel.node.ts    # 메인 노드 구현
│       └── spreadsheet.png  # 노드 아이콘
├── dist/               # 빌드 출력 디렉토리
├── data/               # 런타임 데이터 저장소
├── ref/                # 참조 파일 (Google Sheets 예제)
└── 설정 파일들
```

## 코딩 규칙

### TypeScript 설정
- **Strict Mode**: 모든 strict 옵션 활성화
- **Target**: ES2019
- **Module**: CommonJS
- **Declaration**: 타입 선언 파일 생성
- **Source Maps**: 인라인 소스맵 포함

### n8n 노드 개발 규칙

#### 1. 노드 클래스 구조
```typescript
export class Excel implements INodeType {
    description: INodeTypeDescription = {
        displayName: string,
        name: string,
        icon: string,
        group: string[],
        version: number,
        description: string,
        defaults: object,
        inputs: NodeConnectionType[],
        outputs: NodeConnectionType[],
        properties: INodeProperties[]
    };
}
```

#### 2. 필수 인터페이스 구현
- `INodeType`: 노드의 기본 인터페이스
- `execute()`: 메인 실행 로직
- `loadOptions` 메소드들: 동적 옵션 로딩

#### 3. 파일 명명 규칙
- 노드 파일: `{NodeName}.node.ts`
- 아이콘 파일: 노드와 같은 디렉토리에 `.png` 또는 `.svg`
- 빌드 출력: `dist/{NodeName}/{NodeName}.node.js`

#### 4. 동적 옵션 로딩
- `loadOptionsMethod`: 드롭다운 옵션 동적 로드
- `loadOptionsDependsOn`: 의존성 관계 설정
- 파일 목록, 시트 이름, 컬럼 등 동적 로드

#### 5. 파일 시스템 처리
- 모든 파일은 `data` 디렉토리에 저장
- 디렉토리가 없으면 자동 생성
- 절대 경로 사용 (`process.cwd()` 기반)

#### 6. 에러 처리
```typescript
throw new NodeOperationError(
    this.getNode(),
    'Error message',
    { description: 'Detailed description' }
);
```

## 빌드 및 배포

### 빌드 프로세스
1. `npm run clean`: dist 디렉토리 정리
2. TypeScript 컴파일: `nodes/` → `dist/`
3. Gulp로 아이콘 복사: `.png`, `.svg` 파일

### 스크립트
- `npm run build`: 전체 빌드
- `npm run dev`: 개발 모드 (watch)
- `npm run lint`: ESLint 검사
- `npm test`: Jest 테스트 실행

### 패키지 배포
- `files`: dist 디렉토리만 포함
- `n8n.nodes`: 노드 엔트리 포인트 정의
- `n8n.n8nNodesApiVersion`: API 버전 1

## 특별 고려사항

### CSV 파일 처리
- CSV는 단일 시트로 처리 (Sheet1)
- fast-csv로 파싱 후 xlsx 형식으로 변환
- 저장 시 다시 CSV 형식으로 변환

### Excel 파일 처리
- 다중 시트 지원
- XLSX 라이브러리 직접 사용
- JSON ↔ Sheet 변환 유틸리티 활용

### 바이너리 데이터 처리
- 업로드: 이전 노드의 바이너리 또는 파일 경로
- 다운로드: 바이너리 데이터로 변환하여 전달
- MIME 타입 자동 설정

### 동적 UI 업데이트
- `displayOptions`: 조건부 필드 표시/숨김
- `typeOptions`: 동적 옵션 로딩 설정
- 정규식 패턴으로 파일 확장자 구분

## 테스트 가이드
- Jest 설정: ts-jest preset 사용
- 테스트 제외: node_modules, dist
- 테스트 파일: `*.test.ts`, `*.spec.ts`

## 주의사항
1. 파일 경로는 항상 절대 경로 사용
2. 파일 작업 시 동기 API 사용 (fs.readFileSync 등)
3. 에러 발생 시 NodeOperationError 사용
4. 사용자 입력 검증 필수
5. 한글 문서(README.md) 유지, 코드 주석은 영문