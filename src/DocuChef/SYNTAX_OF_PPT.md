# PowerPoint 템플릿 문법 작성 지침

## 기본 원칙

1. **디자인 중심**: PowerPoint는 디자인이 중요하므로 미리 디자인된 요소에 데이터를 바인딩하는 방식 사용
2. **DollarSignEngine 친화적**: 기존 DollarSignEngine 사용 경험과 최대한 유사하게 유지
3. **간결함**: 직관적이고 간소화된 문법으로 사용 편의성 극대화
4. **디자인 보존**: PPT 디자인 요소의 원래 의도 존중

## 문법 구조

### 1. 값 바인딩 (슬라이드 요소 내)
```
${PropertyName}                // 기본 속성 바인딩
${Object.PropertyName}         // 중첩 속성 바인딩
${Value:FormatSpecifier}       // 포맷 지정자 사용
${Condition ? Value1 : Value2} // 조건부 표현식
${Method()}                    // 메소드 호출
${Array[Index].PropertyName}   // 배열 요소의 속성 바인딩
```

### 2. 특수 함수 (슬라이드 요소 내)
```
${ppt.Image("ImageProperty")}   // 이미지 바인딩
${ppt.Chart("ChartData")}       // 차트 데이터 바인딩
${ppt.Table("TableData")}       // 표 데이터 바인딩
```

### 3. 제어 지시문 (슬라이드 노트에만 배치)
```
#if: Condition, target: "ShapeName", visibleWhenFalse: "AlternateShapeName"  // 조건부 요소 표시/숨김
#foreach: Collection, max: Number, offset: Number                              // 배열 반복 처리 (선택적)
```

## 배열 데이터 처리

PowerPoint는 디자인 중심이므로, 배열 처리도 미리 디자인된 요소에 데이터를 바인딩하는 방식으로 처리합니다:

### 자동 배열 인덱싱

1. **배열 요소 직접 참조**: 슬라이드 요소에서 배열 인덱스를 직접 사용
   ```
   제품 1: ${Products[0].Name} - ${Products[0].Price}
   제품 2: ${Products[1].Name} - ${Products[1].Price}
   제품 3: ${Products[2].Name} - ${Products[2].Price}
   ```

2. **자동 슬라이드 생성**: 엔진이 자동으로 배열 크기에 따라 슬라이드 복제
   - 예: `Products` 배열이 8개 항목을 가지고 있다면:
     - 첫번째 슬라이드: 0,1,2번 항목
     - 두번째 슬라이드: 3,4,5번 항목
     - 세번째 슬라이드: 6,7번 항목

3. **인덱스 오프셋 자동 계산**: 추가된 슬라이드에서 인덱스 자동 조정
   - 첫번째 슬라이드: `Products[0]`, `Products[1]`, `Products[2]`
   - 두번째 슬라이드: 동일한 참조가 `Products[3]`, `Products[4]`, `Products[5]`로 자동 변환
   - 세번째 슬라이드: 동일한 참조가 `Products[6]`, `Products[7]`, 빈 값으로 자동 변환

### 작동 방식

1. **배열 인덱스 패턴 감지**: 엔진이 자동으로 슬라이드 내 `${Array[Index].Property}` 패턴을 분석
2. **슬라이드당 항목 수 계산**: 한 슬라이드에서 가장 높은 인덱스 + 1로 결정
   - 예: 한 슬라이드에 `[0]`, `[1]`, `[2]`가 있다면 슬라이드당 3개 항목
3. **필요한 슬라이드 수 계산**: `총 항목 수 ÷ 슬라이드당 항목 수`로 계산
4. **자동 슬라이드 복제**: 필요한 만큼 원본 슬라이드 복제
5. **인덱스 자동 조정**: 각 복제된 슬라이드에서 인덱스 참조를 자동으로 조정

## foreach 디렉티브 (선택 사항)

`#foreach` 디렉티브는 배열 데이터 처리를 위한 명시적인 방법이지만, **필수가 아닙니다**. 라이브러리는 디자인 요소를 자동으로 분석하여 배열 패턴을 감지하고 처리할 수 있습니다.

### #foreach 문법
```
#foreach: Collection, max: Number, offset: Number
```

- **Collection**: 반복할 배열 또는 컬렉션 (필수)
- **max**: 슬라이드당 최대 항목 수 (선택, 기본값: 자동 감지)
- **offset**: 시작 인덱스 오프셋 (선택, 기본값: 0)

### 자동 감지 vs. 명시적 #foreach

1. **자동 감지 (기본 동작)**
   - 사용자가 `#foreach` 디렉티브를 포함하지 않아도 라이브러리는 배열 패턴을 자동으로 감지
   - 슬라이드 내 `${Array[Index]}` 패턴을 분석하여 필요한 만큼 슬라이드 자동 복제
   - 복제된 슬라이드에서 인덱스 자동 조정

2. **명시적 #foreach (선택 사항)**
   - 보다 세밀한 제어가 필요한 경우 `#foreach` 디렉티브를 사용
   - 슬라이드당 항목 수 명시적 지정 가능
   - 시작 오프셋 지정 가능
   - 중첩된 배열 처리에 유용

### 예시

**자동 감지 방식 (디렉티브 없음)**:
- 슬라이드에 `${Products[0]}`, `${Products[1]}`, `${Products[2]}` 참조가 있으면 자동으로 인식

**명시적 #foreach 사용**:
```
#foreach: Products, max: 3
```
- 슬라이드에 `${Products[0]}`, `${Products[1]}`, `${Products[2]}` 참조를 포함하고
- 슬라이드 노트에 위 디렉티브를 추가

**중첩된 배열에서 명시적 #foreach**:
```
#foreach: Departments
```
- 부서별 슬라이드에서 `${Departments[0].Name}` 참조를 포함하고
- 각 부서의 팀 목록을 표시하는 슬라이드에서는:
```
#foreach: Departments_Teams
```
- 위 디렉티브를 사용하여 현재 부서의 팀 목록에 접근

## 상세 문법 설명

### 1. 값 바인딩 (슬라이드 요소 내)

모든 텍스트 요소에서 DollarSignEngine의 문법을 그대로 사용:

```
제목: ${Report.Title}
날짜: ${DateTime.Now:yyyy-MM-dd}
합계: ${Items.Sum(i => i.Price):C2}
상태: ${Status == "active" ? "활성" : "비활성"}
```

배열 요소 참조:

```
제품: ${Products[0].Name}
가격: ${Products[0].Price:C0}원
설명: ${Products[0].Description}
```

### 2. 특수 함수 (슬라이드 요소 내)

#### 이미지 바인딩
이미지 도형의 텍스트에:
```
${ppt.Image("Company.Logo")}
${ppt.Image("Product.Photo", width: 300, height: 200, preserveAspectRatio: true)}
```

배열 요소의 이미지:
```
${ppt.Image("Products[0].Image")}
```

#### 차트 데이터 바인딩
차트 도형의 텍스트에:
```
${ppt.Chart("SalesData")}
${ppt.Chart("SalesData", series: "Series", categories: "Categories", title: "월별 판매량")}
```

#### 표 데이터 바인딩
표 도형의 텍스트에:
```
${ppt.Table("EmployeeData")}
${ppt.Table("EmployeeData", headers: true, startRow: 1, endRow: 10)}
```

### 3. 제어 지시문 (슬라이드 노트에만 배치)

#### 조건부 지시문
```
#if: Report.HasChart, target: "SalesChart"
#if: Total > 1000, target: "WarningBox", visibleWhenFalse: "SuccessBox"
```
- `target`: 조건부로 표시/숨김 처리할 도형의 이름
- `visibleWhenFalse`: 조건이 거짓일 때 표시할 대체 도형의 이름

#### 반복 지시문 (선택 사항)
```
#foreach: Products, max: 3
#foreach: Departments_Teams, max: 5, offset: 10
```
- 자동 배열 처리가 기본 동작이므로 이 지시문은 옵션입니다
- 기본적으로 라이브러리는 인덱스 패턴으로 배열 처리를 자동으로 수행합니다

## 문법 적용 위치

1. **값 바인딩 & 특수 함수**: 
   - 텍스트 상자, 도형, 표, 차트 등 슬라이드 요소의 텍스트 내용에 배치
   - PowerPoint에서 해당 요소를 선택하고 텍스트 편집 모드에서 입력

2. **제어 지시문**: 
   - 슬라이드 노트에만 배치 (보기 > 노트)
   - 여러 지시문이 필요한 경우 각각 새 줄에 배치

3. **요소 식별**: 
   - PowerPoint에서 도형 선택 → 오른쪽 클릭 → 이름 지정
   - 지정된 이름으로 제어 지시문에서 참조

## 슬라이드 노트 작성 예시

```
# 이 슬라이드는 제품 상세 정보를 표시합니다
#if: Products.Count > 0, target: "ProductsContainer", visibleWhenFalse: "NoProductsMessage"
#foreach: Products, max: 3  # 선택 사항: 명시적으로 지정하지 않아도 자동 처리됨
```

## 예제 시나리오

### 기본 프레젠테이션 슬라이드

**슬라이드 요소 내용:**
- 제목 텍스트 상자: `${Report.Title}`
- 부제목 텍스트 상자: `${Report.Subtitle}`
- 날짜 텍스트 상자: `${Report.Date:yyyy-MM-dd}`
- 로고 이미지: `${ppt.Image("Company.Logo")}`

**슬라이드 노트:**
```
#if: Report.IsConfidential, target: "ConfidentialWatermark"
```

### 제품 목록 슬라이드 (배열 인덱스 사용)

**슬라이드 요소 내용:**
- 제목 텍스트 상자: `${Category.Name} 제품 목록`
- 제품 항목 1: 
  ```
  ${Products[0].Id}. ${Products[0].Name}
  가격: ${Products[0].Price:C0}원
  ```
- 제품 항목 2: 
  ```
  ${Products[1].Id}. ${Products[1].Name}
  가격: ${Products[1].Price:C0}원
  ```
- 제품 항목 3: 
  ```
  ${Products[2].Id}. ${Products[2].Name}
  가격: ${Products[2].Price:C0}원
  ```

**슬라이드 노트:**
```
#if: Products.Count > 0, target: "ProductsContainer"
#foreach: Products, max: 3  # 선택 사항: 라이브러리는 인덱스 패턴을 자동으로 감지합니다
```

**결과:**
- `Products` 배열이 8개 항목을 가지고 있다면:
  - 첫번째 슬라이드: 0,1,2번 항목
  - 두번째 슬라이드: 3,4,5번 항목
  - 세번째 슬라이드: 6,7번 항목

### 데이터 대시보드 슬라이드

**슬라이드 요소 내용:**
- 제목 텍스트 상자: `${Period} 판매 분석`
- 차트 도형 (이름: "SalesChart"): 
  ```
  ${ppt.Chart("SalesData", title: "${Period} 판매 추이")}
  ```
- 표 도형 (이름: "TopProducts"): 
  ```
  ${ppt.Table("TopProducts", headers: true)}
  ```

**슬라이드 노트:**
```
#if: SalesData.Length > 0, target: "SalesChart", visibleWhenFalse: "NoDataMessage"
```

### 부서별 팀원 슬라이드 (배열 인덱스 사용)

**슬라이드 요소 내용:**
- 제목 텍스트 상자: `${Department.Name} 부서`
- 부서장 텍스트 상자: `부서장: ${Department.Manager}`
- 인원수 텍스트 상자: `인원: ${Department.Members.Length}명`
- 팀원 1: `${Department.Members[0].Name} (${Department.Members[0].Position})`
- 팀원 2: `${Department.Members[1].Name} (${Department.Members[1].Position})`
- 팀원 3: `${Department.Members[2].Name} (${Department.Members[2].Position})`
- 팀원 4: `${Department.Members[3].Name} (${Department.Members[3].Position})`
- 팀원 5: `${Department.Members[4].Name} (${Department.Members[4].Position})`
- 팀원 6: `${Department.Members[5].Name} (${Department.Members[5].Position})`

**슬라이드 노트:**
```
#if: Department.Members.Length > 0, target: "MembersContainer"
#foreach: Department.Members, max: 6  # 선택 사항: 인덱스 패턴은 자동으로 감지됩니다
```

**결과:**
- 부서에 15명의 팀원이 있다면:
  - 첫번째 슬라이드: 0~5번 팀원
  - 두번째 슬라이드: 6~11번 팀원
  - 세번째 슬라이드: 12~14번 팀원 (나머지 자리는 빈 값)

## 배열처리 세부 설명

### 중요 개념

1. **디자인 우선**: PowerPoint는 배열 처리에서도 미리 디자인된 요소에 데이터를 바인딩하는 방식 사용
2. **자동 슬라이드 복제**: 배열 데이터가 한 슬라이드에 모두 표시할 수 없는 경우 슬라이드 자동 복제
3. **인덱스 자동 조정**: 복제된 슬라이드에서 인덱스 참조 자동 조정
4. **불필요한 지시문 최소화**: 별도의 반복 지시문 없이 엔진이 배열 패턴 자동 감지

### 작동 방식 상세

1. **인덱스 패턴 분석**: 슬라이드 내 모든 `${Array[Index]}` 참조 스캔
2. **최대 인덱스 파악**: 한 슬라이드 내 가장 높은 인덱스 값 결정
3. **항목수 계산**: 최대 인덱스 + 1 = 한 슬라이드에 표시될 항목 수
4. **필요 슬라이드 계산**: 데이터 크기 ÷ 슬라이드당 항목 수 = 필요한 슬라이드 수
5. **슬라이드 복제**: 원본 슬라이드 디자인을 유지하며 필요한 만큼 복제
6. **인덱스 매핑**: 각 슬라이드에서 인덱스 자동 조정
   - 첫번째 슬라이드: 원본 인덱스 사용
   - 두번째 슬라이드: 원본 인덱스 + 슬라이드당 항목 수
   - 세번째 슬라이드: 원본 인덱스 + (슬라이드당 항목 수 × 2)

### 예시: 부서별 멤버 리스트

**슬라이드 디자인**:
- 제목: `${Department.Name}`
- 멤버 1: `${Department.Members[0].Name}`
- 멤버 2: `${Department.Members[1].Name}`
- 멤버 3: `${Department.Members[2].Name}`

**데이터**: 부서에 8명의 멤버

**결과**:
- 첫번째 슬라이드:
  - 제목: "영업부"
  - 멤버 1: "홍길동" (`Members[0]`)
  - 멤버 2: "김철수" (`Members[1]`)
  - 멤버 3: "이영희" (`Members[2]`)
- 두번째 슬라이드:
  - 제목: "영업부"
  - 멤버 1: "박지성" (`Members[3]`)
  - 멤버 2: "손흥민" (`Members[4]`)
  - 멤버 3: "장미란" (`Members[5]`)
- 세번째 슬라이드:
  - 제목: "영업부"
  - 멤버 1: "김연아" (`Members[6]`)
  - 멤버 2: "이승기" (`Members[7]`)
  - 멤버 3: (빈 값)

### 빈 값 처리

배열 인덱스가 실제 데이터 범위를 초과할 경우:
1. **텍스트 요소**: 빈 문자열("")로 대체
2. **이미지 요소**: 기본 이미지로 대체 또는 숨김
3. **차트/표 요소**: 데이터 없음 상태로 표시 또는 숨김

### 조건부 요소 처리

배열 데이터와 조건부 표시를 결합:
```
#if: Department.Members.Length > ${Index}, target: "Member_${Index}"
```

이 방식으로 해당 인덱스에 멤버가 없는 경우 요소를 숨길 수 있습니다.

### 자동 감지와 #foreach 함께 사용하기

라이브러리는 기본적으로 `${Array[Index]}` 패턴을 자동으로 감지하여 배열 처리하지만, 명시적인 제어가 필요한 경우 `#foreach` 디렉티브를 사용할 수 있습니다. 두 방식은 함께 사용 가능합니다:

1. **자동 감지만 사용**: 대부분의 경우 충분함
   - 인덱스 패턴 분석으로 슬라이드당 항목 수 자동 결정
   - 배열 크기에 따라 슬라이드 자동 복제

2. **#foreach 디렉티브 사용**: 세밀한 제어가 필요한 경우
   - 슬라이드당 항목 수 명시적 지정
   - 시작 오프셋 제어
   - 복잡한 중첩 구조 처리

사용자는 필요에 따라 적절한 방식을 선택할 수 있으며, 대부분의 경우 자동 감지만으로도 충분합니다.