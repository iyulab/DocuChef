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

## 7. 구현 절차 상세 예시

템플릿 엔진이 실제로 어떻게 작동하는지 간단한 예제를 통해 살펴보겠습니다.

### 7.1 시나리오 설정

**템플릿 구성:**
```
슬라이드1:
- ${Items[0].Name}
- ${Items[1].Name}
```

**데이터:**
```
Items = [
 { "Name": "aaa" },
 { "Name": "bbb" }, 
 { "Name": "ccc" }, 
 { "Name": "ddd" }, 
 { "Name": "eee" }, 
]
```

### 7.2 처리 단계별 상세 설명

#### 7.2.1 단계 1: 템플릿 분석 (TemplateAnalyzer)

슬라이드 내의 바인딩 표현식 분석:
- `${Items[0].Name}` - Items 배열의 0번 요소
- `${Items[1].Name}` - Items 배열의 1번 요소

발견된 배열 인덱스 패턴으로 자동 디렉티브 생성:
```
#foreach: Items, max: 2
#range: Items
```

이때:
- `max: 2`는 슬라이드 내 최대 인덱스(1) + 1로 계산
- `#range: Items`는 슬라이드가 Items 컬렉션에 대해 반복됨을 나타냄

#### 7.2.2 단계 2: 슬라이드 계획 생성 (SlidePlanGenerator)

필요한 슬라이드 수 계산:
- Items 컬렉션 크기: 5
- 슬라이드당 항목 수: 2
- 필요한 슬라이드 수: ⌈5 ÷ 2⌉ = 3

슬라이드 인스턴스 계획:
```
슬라이드1: (원본) - 오프셋 0
  - ${Items[0].Name}
  - ${Items[1].Name}

슬라이드2: (복제) - 오프셋 2 
  - ${Items[2].Name}
  - ${Items[3].Name}

슬라이드3: (복제) - 오프셋 4
  - ${Items[4].Name}
  - ${Items[5].Name} // 배열 범위 초과
```

#### 7.2.3 단계 3: 슬라이드 생성 (SlideGenerator)

1. 원본 슬라이드를 복제하여 총 3개의 슬라이드 생성
2. 각 복제 슬라이드의 바인딩 표현식 인덱스 조정:
   - 슬라이드2: 인덱스에 +2 오프셋 적용
   - 슬라이드3: 인덱스에 +4 오프셋 적용

#### 7.2.4 단계 4: 데이터 바인딩 (DataBinder)

바인딩 표현식 해결 및 값 대체:

```
슬라이드1: (원본)
- aaa  // Items[0].Name
- bbb  // Items[1].Name

슬라이드2: (복제)
- ccc  // Items[2].Name
- ddd  // Items[3].Name

슬라이드3: (복제)
- eee  // Items[4].Name
- (빈 값)  // Items[5].Name - 배열 범위 초과로 빈 값 표시
```

### 7.3 주요 특징 설명

- **인덱스 자동 조정**: 복제된 슬라이드의 인덱스가 자동으로 조정됨
- **범위 초과 처리**: 배열 범위를 초과하는 인덱스는 빈 값으로 처리됨
- **디렉티브 자동 생성**: 슬라이드 내 표현식 패턴에서 디렉티브 자동 추론
- **슬라이드 수 최적화**: 데이터 크기에 맞게 필요한 슬라이드만 생성

이 예시는 템플릿 엔진의 핵심 처리 흐름을 보여주며, 실제 구현에서는 더 복잡한 중첩 컬렉션, 컨텍스트 연산자, 조건부 표현식 등을 지원합니다.

## 8. 컨텍스트 연산자('>') 활용 중첩 구조 처리 예시

중첩된 데이터 구조를 처리할 때 컨텍스트 연산자('>')가 어떻게 처리되는지 상세히 살펴보겠습니다.

### 8.1 시나리오 설정

**템플릿 구성:**
```
슬라이드1 (카테고리): 
- 카테고리: ${Categories[0].Name}
#foreach: Categories
#range: begin, Categories

슬라이드2 (제품):
- 제품1: ${Categories>Products[0].Name} - ${Categories>Products[0].Price}
- 제품2: ${Categories>Products[1].Name} - ${Categories>Products[1].Price}
#foreach: Categories>Products, max: 2
#range: end, Categories
```

**데이터:**
```
Categories = [
  { 
    "Name": "전자기기", 
    "Products": [
      { "Name": "스마트폰", "Price": 1000000 },
      { "Name": "태블릿", "Price": 800000 },
      { "Name": "노트북", "Price": 1500000 }
    ]
  },
  { 
    "Name": "가구", 
    "Products": [
      { "Name": "소파", "Price": 500000 },
      { "Name": "침대", "Price": 700000 },
      { "Name": "책상", "Price": 300000 },
      { "Name": "의자", "Price": 150000 }
    ]
  }
]
```

### 8.2 컨텍스트 연산자 처리 단계별 상세 설명

#### 8.2.1 단계 1: 템플릿 분석 (TemplateAnalyzer)

**슬라이드1 분석:**
- `${Categories[0].Name}` - 배열 요소 참조
- 생성된 디렉티브: `#foreach: Categories`, `#range: begin, Categories`
- 슬라이드 유형: Source (범위 시작)

**슬라이드2 분석:**
- `${Categories>Products[0].Name}` - 컨텍스트 연산자를 통한 현재 카테고리의 Products 참조
- `${Categories>Products[1].Name}` - 컨텍스트 연산자를 통한 현재 카테고리의 Products 참조
- 생성된 디렉티브: `#foreach: Categories>Products, max: 2`, `#range: end, Categories`
- 슬라이드 유형: End (범위 종료)
- 컨텍스트 경로: `Categories>Products` (현재 카테고리의 Products를 참조)

#### 8.2.2 단계 2: 슬라이드 계획 생성 (SlidePlanGenerator)

**컨텍스트 체인 분석:**
1. 최상위 컬렉션: `Categories` (2개 항목)
2. 각 카테고리의 하위 컬렉션: `Products`
   - 첫 번째 카테고리: 3개 제품
   - 두 번째 카테고리: 4개 제품

**슬라이드 인스턴스 계획:**
```
카테고리1 슬라이드:
- 카테고리: 전자기기
컨텍스트: Categories[0]

카테고리1-제품 슬라이드1:
- 제품1: 스마트폰 - 1000000
- 제품2: 태블릿 - 800000
컨텍스트: Categories[0]>Products[0,1]

카테고리1-제품 슬라이드2:
- 제품1: 노트북 - 1500000
- 제품2: (빈 값)
컨텍스트: Categories[0]>Products[2,3] (인덱스 3은 범위 초과)

카테고리2 슬라이드:
- 카테고리: 가구
컨텍스트: Categories[1]

카테고리2-제품 슬라이드1:
- 제품1: 소파 - 500000
- 제품2: 침대 - 700000
컨텍스트: Categories[1]>Products[0,1]

카테고리2-제품 슬라이드2:
- 제품1: 책상 - 300000
- 제품2: 의자 - 150000
컨텍스트: Categories[1]>Products[2,3]
```

**필요한 슬라이드 총 개수:**
- 카테고리 슬라이드: 2개
- 제품 슬라이드: (⌈3÷2⌉) + (⌈4÷2⌉) = 2 + 2 = 4개
- 총 슬라이드: 6개

#### 8.2.3 단계 3: 슬라이드 생성 및 표현식 조정 (SlideGenerator)

**컨텍스트 연산자 처리:**
슬라이드 복제 시 각 인스턴스의 컨텍스트 정보 설정:

1. **카테고리1-제품 슬라이드1:**
   원본: `${Categories>Products[0].Name}` → 처리: 현재 컨텍스트 `Categories[0]`에서 `Products[0]` 참조
   결과: `스마트폰` (Categories[0].Products[0].Name)

2. **카테고리1-제품 슬라이드2:**
   - 표현식 조정: `${Categories>Products[0].Name}` → `${Categories>Products[2].Name}`
   - 오프셋 +2 적용 (슬라이드당 항목 수 2개 * 인스턴스 순번 1)
   - 컨텍스트: Categories[0]

3. **카테고리2-제품 슬라이드1:**
   - 표현식 유지: `${Categories>Products[0].Name}`
   - 컨텍스트: Categories[1]로 변경
   - 결과: `소파` (Categories[1].Products[0].Name)

4. **카테고리2-제품 슬라이드2:**
   - 표현식 조정: `${Categories>Products[0].Name}` → `${Categories>Products[2].Name}`
   - 오프셋 +2 적용
   - 컨텍스트: Categories[1]
   - 결과: `책상` (Categories[1].Products[2].Name)

#### 8.2.4 단계 4: 데이터 바인딩 (DataBinder)

**컨텍스트 연산자 해석:**
1. 컨텍스트 체인 분석 및 현재 컨텍스트 식별
2. 표현식의 컨텍스트 연산자('>') 처리
   - 컨텍스트 연산자 이전 부분 → 컨텍스트 식별자
   - 컨텍스트 연산자 이후 부분 → 현재 컨텍스트에서의 상대 경로

**최종 바인딩 결과:**
```
카테고리1 슬라이드:
- 카테고리: 전자기기

카테고리1-제품 슬라이드1:
- 제품1: 스마트폰 - 1000000
- 제품2: 태블릿 - 800000

카테고리1-제품 슬라이드2:
- 제품1: 노트북 - 1500000
- 제품2: (빈 값)

카테고리2 슬라이드:
- 카테고리: 가구

카테고리2-제품 슬라이드1:
- 제품1: 소파 - 500000
- 제품2: 침대 - 700000

카테고리2-제품 슬라이드2:
- 제품1: 책상 - 300000
- 제품2: 의자 - 150000
```

### 8.3 컨텍스트 연산자('>') 처리 핵심 원리

1. **컨텍스트 인식:**
   - '>' 왼쪽: 컨텍스트 식별자 (어떤 컬렉션의 현재 항목을 참조하는지)
   - '>' 오른쪽: 컨텍스트 내의 상대 경로

2. **다중 슬라이드 생성 규칙:**
   - 각 컨텍스트 레벨마다 필요한 슬라이드 계산
   - 중첩 컬렉션은 부모 컨텍스트를 기준으로 처리

3. **인덱스 조정 메커니즘:**
   - 부모 컨텍스트 변경: 상위 컬렉션의 다른 항목으로 이동
   - 인덱스 오프셋 적용: 동일 컨텍스트 내에서 인덱스 조정

4. **경로 해석 순서:**
   - 컨텍스트 식별 → 현재 컨텍스트 경로 결정 → 상대 경로 적용 → 값 해석

이 예시는 컨텍스트 연산자('>')를 활용한 중첩 데이터 구조 처리 방식을 보여주며, 복잡한 계층 구조에서도 직관적인 템플릿 작성이 가능하게 합니다.

## 9. 바인딩이 없는 슬라이드 위치 유지 처리

템플릿 프레젠테이션에는 데이터 바인딩이 없는 일반 슬라이드와 바인딩이 있는 동적 슬라이드가 혼합되어 있는 경우가 많습니다. 이러한 경우 각 슬라이드의 원래 위치와 순서를 유지하는 것이 중요합니다.

### 9.1 시나리오 설정

**프레젠테이션 구성:**
```
슬라이드1: 표지 (바인딩 없음)
슬라이드2: 목차 (바인딩 없음)
슬라이드3: 제품 소개 (바인딩 있음)
  - ${Products[0].Name}
  - ${Products[1].Name}
  #foreach: Products, max: 2
슬라이드4: 중간 설명 (바인딩 없음)
슬라이드5: 카테고리 상세 (바인딩 있음)
  - ${Categories[0].Name}
  #foreach: Categories
슬라이드6: 마무리 (바인딩 없음)
```

**데이터:**
```
Products = [
  { "Name": "제품A" },
  { "Name": "제품B" },
  { "Name": "제품C" },
  { "Name": "제품D" },
  { "Name": "제품E" }
]

Categories = [
  { "Name": "카테고리1" },
  { "Name": "카테고리2" },
  { "Name": "카테고리3" }
]
```

### 9.2 위치 유지 처리 단계별 설명

#### 9.2.1 단계 1: 슬라이드 유형 분류 (TemplateAnalyzer)

슬라이드 분석 시 각 슬라이드를 다음과 같이 분류합니다:

1. **정적 슬라이드 (Static)**: 바인딩 표현식이 없는 슬라이드
   - 슬라이드1, 슬라이드2, 슬라이드4, 슬라이드6

2. **동적 슬라이드 (Dynamic)**: 바인딩 표현식이 있는 슬라이드
   - 슬라이드3 (Products 컬렉션 바인딩)
   - 슬라이드5 (Categories 컬렉션 바인딩)

각 슬라이드의 원래 순서와 위치 정보 기록:
```
SlideInfo[0]: { SlideId: 1, Type: Static, Position: 0 }
SlideInfo[1]: { SlideId: 2, Type: Static, Position: 1 }
SlideInfo[2]: { SlideId: 3, Type: Source, Position: 2, CollectionName: "Products" }
SlideInfo[3]: { SlideId: 4, Type: Static, Position: 3 }
SlideInfo[4]: { SlideId: 5, Type: Source, Position: 4, CollectionName: "Categories" }
SlideInfo[5]: { SlideId: 6, Type: Static, Position: 5 }
```

#### 9.2.2 단계 2: 위치 기준점 설정 (SlidePlanGenerator)

정적 슬라이드의 위치를 기준점으로 사용하여 동적 슬라이드 생성 위치 결정:

1. **기준점 맵 구성**:
   - 정적 슬라이드의 위치를 키로 하는 맵 생성
   - 각 동적 슬라이드 영역의 시작과 끝 위치 식별

```
AnchorPoints = {
  0: { SlideId: 1 },  // 첫 번째 정적 슬라이드 (표지)
  1: { SlideId: 2 },  // 두 번째 정적 슬라이드 (목차)
  3: { SlideId: 4 },  // 세 번째 정적 슬라이드 (중간 설명)
  5: { SlideId: 6 }   // 네 번째 정적 슬라이드 (마무리)
}

DynamicRanges = {
  "Products": { Start: 2, End: 2 },  // 제품 슬라이드 영역
  "Categories": { Start: 4, End: 4 }  // 카테고리 슬라이드 영역
}
```

#### 9.2.3 단계 3: 슬라이드 계획 생성 (SlidePlanGenerator)

정적 슬라이드의 위치를 유지하면서 동적 슬라이드를 삽입하는 계획 생성:

1. **Products 컬렉션 처리**:
   - 필요 슬라이드 수: ⌈5 ÷ 2⌉ = 3
   - 위치: 원래 위치 (2) 유지하고 연속 배치

2. **Categories 컬렉션 처리**:
   - 필요 슬라이드 수: 3
   - 위치: 원래 위치 (4) 유지하고 연속 배치

**슬라이드 계획 (오프셋 고려):**
```
1. 슬라이드1: 표지 (원본 유지)
2. 슬라이드2: 목차 (원본 유지)
3. 슬라이드3: 제품 소개 - 원본 (Products[0,1])
4. 슬라이드3': 제품 소개 - 복제 1 (Products[2,3])
5. 슬라이드3": 제품 소개 - 복제 2 (Products[4])
6. 슬라이드4: 중간 설명 (원본 유지)
7. 슬라이드5: 카테고리 상세 - 원본 (Categories[0])
8. 슬라이드5': 카테고리 상세 - 복제 1 (Categories[1])
9. 슬라이드5": 카테고리 상세 - 복제 2 (Categories[2])
10. 슬라이드6: 마무리 (원본 유지)
```

#### 9.2.4 단계 4: 슬라이드 위치 조정 (SlideGenerator)

최종 프레젠테이션에서 슬라이드 위치 계산 및 조정:

1. **정적 슬라이드**: 원래 템플릿의 슬라이드 ID 참조
2. **동적 슬라이드**: 원본 슬라이드 ID에 새 슬라이드 ID 할당
3. **슬라이드 순서**: 계획에 따라 모든 슬라이드 배치

**최종 슬라이드 순서 (SlideId 매핑):**
```
PresentationSlides = [
  { PresentationIndex: 0, TemplateSlideId: 1 },  // 표지
  { PresentationIndex: 1, TemplateSlideId: 2 },  // 목차
  { PresentationIndex: 2, TemplateSlideId: 3 },  // 제품 소개 (Products[0,1])
  { PresentationIndex: 3, TemplateSlideId: 3 },  // 제품 소개 (Products[2,3])
  { PresentationIndex: 4, TemplateSlideId: 3 },  // 제품 소개 (Products[4])
  { PresentationIndex: 5, TemplateSlideId: 4 },  // 중간 설명
  { PresentationIndex: 6, TemplateSlideId: 5 },  // 카테고리 상세 (Categories[0])
  { PresentationIndex: 7, TemplateSlideId: 5 },  // 카테고리 상세 (Categories[1])
  { PresentationIndex: 8, TemplateSlideId: 5 },  // 카테고리 상세 (Categories[2])
  { PresentationIndex: 9, TemplateSlideId: 6 }   // 마무리
]
```

### 9.3 위치 유지 처리 핵심 원리

1. **슬라이드 분류 메커니즘**:
   - 바인딩 표현식 유무에 따른 정적/동적 슬라이드 분류
   - 각 슬라이드의 원래 위치(Position) 정보 유지

2. **상대적 위치 보존 알고리즘**:
   - 정적 슬라이드는 항상 원래 상대적 위치에 배치
   - 동적 슬라이드는 원래 위치를 시작점으로 연속 배치
   - 기준점(Anchor) 기반의 슬라이드 배치 계산

3. **다중 컬렉션 처리 규칙**:
   - 각 동적 슬라이드 영역을 독립적으로 확장
   - 다른 영역에 영향을 주지 않도록 위치 오프셋 적용

4. **슬라이드 ID 관리**:
   - 원본 템플릿의 슬라이드 ID를 참조 정보로 유지
   - 새 프레젠테이션에서 적절한 순서로 슬라이드 배치

이 접근 방식을 통해 바인딩이 없는 정적 슬라이드(표지, 목차, 설명, 마무리 등)의 위치를 그대로 유지하면서, 동적 데이터에 따라 필요한 슬라이드를 자동 생성하고 배치할 수 있습니다.