# ExcelReportManager

## 소개
ExcelReportManager는 엑셀 리포트를 생성, 수정, 관리하기 위한 유틸리티 및 전략 기반 모듈

---

## 목적
ExcelReportManager는 프로젝트에서 반복적인 엑셀 리포트 생성 및 관리 작업을 보다 효율화하기 위해 설계되었습니다.
이를 통해 코드를 단순화하고 유지보수를 용이하게 할 수 있습니다.

---

## 주요 구성 요소
- **ExcelFactory**: 엑셀 파일 생성 도구.
- **ExcelFileUtils/ExcelSheetUtils**: 파일 및 시트 작업 유틸리티.
- **ExcelReportProcessor**: 리포트 생성/처리 핵심 클래스.
- **ExcelReportStrategy**: 커스텀 엑셀 리포트 생성을 위한 전략 패턴 인터페이스.
- **NaverReportStrategy**: 네이버 광고 리포트 생성을 위한 구현체.


### NaverReportStrategy

`NaverReportStrategy`는 Excel 리포트를 생성하기 위한 전략 클래스입니다.  
해당 클래스의 구현 로직 **`fillReport` 메서드**는 각자의 요구 사항에 따라 **직접 커스텀**하여 사용해야 합니다.

### 참고 사항
- `NaverReportStrategy`는 `ExcelReportStrategy` 인터페이스를 구현합니다.
- `fillReport(Workbook workbook, ReportData reportData)` 메서드를 커스텀하여, 원하는 리포트 생성 로직을 추가하세요.
- 구체적인 비즈니스 로직은 사용자에 따라 다를 수 있으며, 이를 기반으로 리포트를 동적으로 생성할 수 있습니다.

### 예시
```java
public class NaverReportStrategy implements ExcelReportStrategy {

    @Override
    public void fillReport(Workbook workbook, ReportData reportData) {
        // 사용자의 요구 사항에 맞게 리포트 생성 로직을 정의하세요.
        // 예: 특정 리포트 데이터를 Excel 시트에 추가하거나 서식을 지정하는 로직 구현
    }

}
```

위 내용을 참고하여 필요에 맞는 구현을 자유롭게 구성하세요.
---

## 사용 예시
```java
// 프로젝트에 의존성 추가
@Configuration
public class ExcelManagerConfiguration {

    @Value("${excel.template-dir}")
    private String templateDir;

    @Bean
    public ExcelReportProcessor excelReportProcessor() {
        Map<ReportType, ExcelReportStrategy> strategies = Map.of(
                ReportType.NAVER, new NaverReportStrategy()
        );
        return new ExcelReportProcessor(strategies, Path.of(templateDir));
    }

}
```

---

## 요구사항
- **Java 버전**: 17 이상
- **사용 라이브러리**: Apache POI, Lombok