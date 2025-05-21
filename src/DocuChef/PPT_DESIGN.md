# PowerPoint 템플릿 엔진 설계 문서

## 1. 시스템 아키텍처

### 1.1 핵심 컴포넌트

PowerPoint 템플릿 엔진은 다음 4단계 처리 흐름을 담당하는 컴포넌트로 구성됩니다:

- **TemplateAnalyzer**: 템플릿 분석 및 SlideInfo 생성
- **SlidePlanGenerator**: 슬라이드 플랜 결정
- **SlideGenerator**: 슬라이드 생성 및 expression 재조정
- **DataBinder**: 데이터 바인딩

### 1.2 핵심 모델

- **SlideInfo**: 슬라이드 ID, 유형(General/Source/Cloned), 디렉티브 목록, 바인딩 표현식 정보
- **SlidePlan**: 생성할 슬라이드 인스턴스 목록
- **SlideInstance**: 원본 슬라이드 ID, 유형, 위치, 컨텍스트, 인덱스 오프셋
- **Directive**: 지시문 유형(foreach/range/alias), 컬렉션 경로, 옵션
- **BindingExpression**: 원본 표현식, 데이터 경로, 배열 인덱스, 형식 지정자

## 2. 구현 플로우

### 2.1 단계 1: 템플릿 분석 - SlideInfo 생성

템플릿을 스캔하여 SlideInfo 객체를 생성합니다:

1. **슬라이드 노트 분석**
   - `#foreach:`, `#range:`, `#alias:` 디렉티브 파싱
   - 디렉티브가 없어도 문제없음 (자동 생성)

2. **바인딩 표현식 스캔**
   - 슬라이드 텍스트 요소에서 `${...}` 패턴 추출
   - 배열 인덱스 패턴 `[n]` 및 컨텍스트 참조 `>` 패턴 분석

3. **자동 디렉티브 생성**
   - 명시적 디렉티브가 없는 경우 표현식 분석으로 자동 생성
   - 배열 패턴에서 슬라이드당 최대 항목 수 결정

### 2.2 단계 2: 슬라이드 플랜 결정

템플릿 분석 결과와 Data를 통해 총 슬라이드 계획을 생성합니다:

1. **필요한 슬라이드 수 계산**
   - 각 컬렉션 크기 ÷ 슬라이드당 항목 수로 계산
   - 올림 처리로 모든 데이터가 표시되도록 보장

2. **컨텍스트 체인 구성**
   - 중첩된 컬렉션 경로 분석
   - 각 슬라이드 인스턴스의 컨텍스트 정보 설정

3. **SlideInstance 생성**
   - 각 필요한 슬라이드에 대한 인스턴스 정보 생성
   - 원본 슬라이드 참조, 생성 위치, 인덱스 오프셋 설정

### 2.3 단계 3: 슬라이드 생성

플랜에 따라 슬라이드를 생성하고 expression의 index를 오프셋 적용하여 재조정합니다:

1. **슬라이드 복제**
   - 원본 슬라이드를 필요한 만큼 복제
   - 지정된 위치에 배치

2. **인덱스 오프셋 적용**
   - 배열 인덱스 패턴에 오프셋 적용
   - 예: `Items[0]` → `Items[3]` (오프셋 +3)

3. **표현식 재조정**
   - 복제된 슬라이드의 모든 바인딩 표현식 업데이트

### 2.4 단계 4: 데이터 바인딩

생성된 슬라이드에 실제 데이터를 바인딩합니다:

1. **표현식 해결**
   - 바인딩 표현식을 데이터 객체와 컨텍스트로 해결

2. **값 대체**
   - `${...}` 패턴을 해결된 값으로 대체
   - 형식 지정자 적용

## 3. 디렉티브 구문

### 3.1 통합된 #range 디렉티브

슬라이드가 복제되는 기준을 제공합니다.

- **`#range: SourceName`**: 단일 슬라이드 범위 (기본)
- **`#range: begin, SourceName`**: 범위 시작
- **`#range: end, SourceName`**: 범위 종료

### 3.2 디렉티브 종류

1. **#foreach: Collection, max: N, offset: M**
   - 컬렉션 반복 처리
   - max, offset은 선택사항 (자동 감지)

2. **#range: [begin|end,] SourceName**
   - 범위 정의 (단일 슬라이드 또는 다중 슬라이드)

3. **#alias: SourcePath as AliasName**
   - 컬렉션 경로에 별칭 부여
   - 긴 컨텍스트 경로를 짧은 이름으로 사용

### 3.3 예시

```
슬라이드1: 
  - ${Categories[0].Name}
  #foreach: Categories
  #range: begin, Categories

슬라이드2: 
  - ${Items[0].Name}
  - ${Items[1].Name}
  #foreach: Categories>Products
  #range: Categories>Products
  #range: end, Categories
  #alias: Categories>Products as Items
```

## 4. 컨텍스트 참조('>' 연산자) 처리

### 4.1 컨텍스트 체인 관리

1. **경로 분할**: '>' 연산자 기준으로 경로 세그먼트 분할
2. **컨텍스트 해결**: 각 세그먼트를 현재 컨텍스트 내에서 해결
3. **체인 업데이트**: 중첩 처리 시 컨텍스트 체인 동적 업데이트

### 4.2 별칭(#alias) 처리

1. **경로 단순화**: 긴 컨텍스트 경로를 짧은 별칭으로 매핑
2. **표현식 변환**: 별칭을 사용한 표현식을 실제 경로로 변환
3. **스코프 관리**: 별칭의 유효 범위 관리

## 5. 중첩 구조 처리 예시

### 5.1 카테고리 > 제품 구조

**템플릿 구성:**
```
슬라이드1 (카테고리):
  - ${Categories[0].Name}
  #foreach: Categories
  #range: begin, Categories

슬라이드2 (제품 목록):
  - ${Items[0].Name} - ${Items[0].Price}
  - ${Items[1].Name} - ${Items[1].Price}
  #foreach: Categories>Products
  #range: Categories>Products
  #range: end, Categories
  #alias: Categories>Products as Items
```

**처리 결과:**
1. Categories 컬렉션의 각 항목에 대해:
   - 카테고리 슬라이드 1개
   - 해당 카테고리의 Products에 대한 제품 목록 슬라이드 N개

### 5.2 4단계 중첩: Reports > Departments > Teams > Members

**슬라이드 계층:**
- 보고서 타이틀 슬라이드: `Reports[0]`
- 부서 개요 슬라이드: `Reports>Departments[0]`, `Reports>Departments[1]`...
- 팀 상세 슬라이드: `Reports>Departments>Teams[0]`, `Reports>Departments>Teams[1]`...
- 팀원 목록 슬라이드: 각 팀의 Members 배열 크기에 따라 복수 생성

**슬라이드 계산:**
```
총 슬라이드 = Reports.length + 
             (Reports.length × Departments.length) + 
             Σ(각 부서별 Teams.length) + 
             Σ(각 팀별 ⌈Members.length ÷ 멤버슬라이드당항목수⌉)
```

## 6. 자동 디렉티브 생성

### 6.1 패턴 분석

1. **배열 인덱스 패턴**: `${Array[0]}`, `${Array[1]}` 등에서 슬라이드당 항목 수 감지
2. **컨텍스트 참조 패턴**: `${Parent>Child[0]}` 등에서 컨텍스트 체인 감지
3. **최대 인덱스 계산**: 슬라이드 내 가장 높은 배열 인덱스를 기준으로 항목 수 결정

### 6.2 암시적 디렉티브 생성

명시적 디렉티브가 없는 경우:
1. 배열 패턴 분석으로 `#foreach` 디렉티브 생성
2. 컨텍스트 참조 패턴으로 `#range` 디렉티브 생성