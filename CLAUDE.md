# CLAUDE.md — my-dashboard

이 파일은 AI 어시스턴트(Claude 등)가 이 레포지토리에서 작업할 때 참고하는 가이드입니다.

## 프로젝트 개요

`my-dashboard`는 개인용 대시보드 웹 애플리케이션 프로젝트입니다.
현재 초기 단계이며, 아래 가이드라인을 따라 개발을 진행합니다.

## 레포지토리 구조

```
my-dashboard/
├── CLAUDE.md          # AI 어시스턴트 가이드 (이 파일)
└── (초기 상태 — 아래 권장 구조 참고)
```

### 권장 프로젝트 구조 (향후 확장 시)

```
my-dashboard/
├── public/            # 정적 파일 (favicon, index.html 등)
├── src/
│   ├── components/    # 재사용 가능한 UI 컴포넌트
│   ├── pages/         # 페이지 단위 컴포넌트
│   ├── hooks/         # 커스텀 React 훅
│   ├── utils/         # 유틸리티 함수
│   ├── services/      # API 호출 및 외부 서비스 연동
│   ├── styles/        # 전역 스타일 및 테마
│   ├── types/         # TypeScript 타입 정의
│   └── App.tsx        # 메인 애플리케이션 엔트리
├── tests/             # 테스트 파일
├── package.json       # 의존성 및 스크립트
├── tsconfig.json      # TypeScript 설정
├── CLAUDE.md          # AI 어시스턴트 가이드
└── README.md          # 프로젝트 설명
```

## 개발 환경

### 기술 스택 (권장)

- **프레임워크**: React + TypeScript
- **빌드 도구**: Vite
- **스타일링**: CSS Modules 또는 Tailwind CSS
- **상태 관리**: React Context API 또는 Zustand
- **테스트**: Vitest + React Testing Library
- **린팅**: ESLint + Prettier

### 주요 명령어

프로젝트 초기화 후 사용할 명령어:

```bash
npm install          # 의존성 설치
npm run dev          # 개발 서버 실행
npm run build        # 프로덕션 빌드
npm run test         # 테스트 실행
npm run lint         # 린트 검사
npm run format       # 코드 포맷팅
```

## 코딩 컨벤션

### 일반 규칙

- 모든 코드와 커밋 메시지는 **영어**로 작성
- 주석과 문서는 **한국어** 허용
- 파일명은 **PascalCase** (컴포넌트) 또는 **camelCase** (유틸리티)
- 디렉터리명은 **kebab-case** 또는 **camelCase**

### TypeScript

- `any` 타입 사용 금지 — 구체적인 타입을 정의할 것
- 인터페이스는 `I` 접두사 없이 명확한 이름 사용 (예: `User`, `DashboardConfig`)
- 타입 정의는 `src/types/` 디렉터리에 모아서 관리

### React 컴포넌트

- 함수형 컴포넌트만 사용 (클래스형 컴포넌트 사용 금지)
- 컴포넌트 파일 하나에 하나의 export default 컴포넌트
- Props 타입은 컴포넌트 파일 내에서 정의
- 커스텀 훅은 `use` 접두사 사용 (예: `useDashboardData`)

### 스타일링

- 인라인 스타일 최소화
- 글로벌 스타일 대신 모듈화된 스타일 우선 사용
- 반응형 디자인 적용 (모바일 우선)

## Git 워크플로우

### 브랜치 전략

- `main`: 안정적인 프로덕션 코드
- `feature/*`: 새 기능 개발
- `fix/*`: 버그 수정
- `claude/*`: AI 어시스턴트 작업용

### 커밋 메시지 규칙

[Conventional Commits](https://www.conventionalcommits.org/) 형식을 따릅니다:

```
<type>: <description>

[optional body]
```

타입 종류:
- `feat`: 새 기능 추가
- `fix`: 버그 수정
- `docs`: 문서 변경
- `style`: 코드 포맷팅 (기능 변경 없음)
- `refactor`: 리팩터링
- `test`: 테스트 추가/수정
- `chore`: 빌드/설정 변경

### PR 규칙

- PR 제목은 간결하게 (70자 이내)
- 변경 사항 요약과 테스트 계획 포함
- 관련 이슈 번호 연결

## AI 어시스턴트를 위한 지침

### 작업 시 주의 사항

1. **코드 읽기 우선**: 수정 전에 반드시 기존 코드를 먼저 읽을 것
2. **최소 변경 원칙**: 요청된 작업에 필요한 최소한의 변경만 수행
3. **보안 주의**: 하드코딩된 시크릿, API 키 등을 커밋하지 않을 것
4. **테스트 작성**: 새 기능 추가 시 관련 테스트도 함께 작성
5. **기존 패턴 준수**: 프로젝트에 이미 확립된 패턴과 구조를 따를 것

### 금지 사항

- `.env` 파일이나 시크릿이 포함된 파일을 커밋하지 않을 것
- `node_modules/` 디렉터리를 커밋하지 않을 것
- `--force` 플래그로 git push 하지 않을 것
- 사용자 확인 없이 파괴적인 작업 수행하지 않을 것
- 불필요한 의존성 추가하지 않을 것
