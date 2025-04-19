# 📘 ExcelProcessor

POI 기반 엑셀 리포트 유틸리티 라이브러리  
템플릿 업로드, 시트별 데이터 삽입, 멀티시트 리포트 생성 기능 제공

---

## 📦 설치 (JitPack)

```gradle
repositories {
    maven { url 'https://jitpack.io' }
}

dependencies {
    implementation 'com.github.사용자명:Excel-Library:버전'
}
```

> `사용자명`과 `버전`은 JitPack에서 배포한 주소를 확인하세요.

---

## 🚀 주요 기능

| 기능 | 설명 |
|------|------|
| 템플릿 저장 | 업로드한 엑셀 파일을 서버에 저장 |
| 시트 목록 조회 | 템플릿에 포함된 시트 이름 리스트 반환 |
| 단일 시트 리포트 | 특정 시트에 데이터 삽입 후 결과 반환 |
| 멀티 시트 리포트 | 여러 시트에 각각 데이터 삽입 가능 |
| 열 너비 자동 조절 | 시트 내 데이터 길이에 따라 자동 맞춤 적용 |

---

## 🧩 클래스 구성

- `ExcelService` : 라이브러리 통합 진입점
- `TemplateManager` : 템플릿 저장소 관리
- `ExcelFileProcessor` : 엑셀 읽기/쓰기 처리
- `ExcelSheetWriter` : 시트별 데이터 삽입 및 기본 설정
- `SheetDataRequest` : 멀티 시트용 데이터 모델

---

## ✅ 사용 예제

### 1. 템플릿 저장

```java
ExcelService service = new ExcelService(new File("excel_templates"));
try (InputStream in = new FileInputStream("template.xlsx")) {
    service.saveTemplate("template.xlsx", in);
}
```

### 2. 시트 이름 조회

```java
List<String> sheets = service.getSheetNames("template.xlsx");
```

### 3. 단일 시트 리포트 생성

```java
List<Map<String, String>> data = List.of(
    Map.of("Name", "Alice", "Score", "95"),
    Map.of("Name", "Bob", "Score", "88")
);

ByteArrayOutputStream out = service.generateReport("template.xlsx", "ReportSheet", data);
Files.write(Paths.get("report.xlsx"), out.toByteArray());
```

### 4. 멀티 시트 리포트 생성

```java
List<SheetDataRequest> sheets = List.of(
    new SheetDataRequest("Sales", List.of(
        Map.of("Product", "Laptop", "Quantity", "10"),
        Map.of("Product", "Mouse", "Quantity", "25")
    )),
    new SheetDataRequest("Summary", List.of(
        Map.of("Total", "35", "Date", "2025-04-19")
    ))
);

ByteArrayOutputStream out = service.generateMultiSheetReport("template.xlsx", sheets);
Files.write(Paths.get("multi_report.xlsx"), out.toByteArray());
```

---

## 🌐 Spring REST API 연동 예시

```java
@RestController
@RequestMapping("/api/excel")
public class ExcelController {
    private final ExcelService service;

    public ExcelController(ExcelService service) {
        this.service = service;
    }

    @PostMapping("/upload")
    public ResponseEntity<String> upload(@RequestParam MultipartFile file) {
        try (InputStream in = file.getInputStream()) {
            service.saveTemplate(file.getOriginalFilename(), in);
            return ResponseEntity.ok("Uploaded");
        } catch (IOException e) {
            return ResponseEntity.status(500).body("Error");
        }
    }
}
```

---

## 📂 템플릿 저장소 경로

라이브러리 생성 시 지정한 `File` 디렉토리에 템플릿이 저장됩니다:

```java
new ExcelService(new File("excel_templates"))
```

---

## 📌 의존성

- Java 17+
- Apache POI

```gradle
implementation 'org.apache.poi:poi:5.2.5'
implementation 'org.apache.poi:poi-ooxml:5.2.5'
```

---

## 🛠 To-Do

- 셀 스타일 지정 기능
- 셀 위치 지정 삽입 (row, col 지정)
- 동적 시트 생성 옵션

---

## 📄 라이선스

MIT License
