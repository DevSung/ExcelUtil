package com.devsung.excel.processor;


import com.devsung.excel.core.ExcelFactory;
import com.devsung.excel.model.ReportData;
import com.devsung.excel.model.ReportType;
import com.devsung.excel.strategy.ExcelReportStrategy;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

public class ExcelReportProcessor {

    private final Map<ReportType, ExcelReportStrategy> strategyMap;
    private final Path templateDir; // 기본 템플릿 디렉토리

    public ExcelReportProcessor(Map<ReportType, ExcelReportStrategy> strategyMap, Path templateDir) {
        this.strategyMap = strategyMap;
        this.templateDir = templateDir;
    }

    /**
     * 템플릿 파일 경로를 생성합니다.
     *
     * @param templateName 템플릿 파일 이름
     * @param paths        템플릿 파일 경로
     * @return 템플릿 파일 경로
     */
    public Path resolveTemplatePath(String templateName, String... paths) {
        if (paths == null || paths.length == 0) {
            return templateDir.resolve(templateName);
        }
        Path result = templateDir;
        for (String path : paths) {
            result = result.resolve(path);
        }
        return result.resolve(templateName);

    }

    /**
     * 리포트를 생성 (기본 템플릿 디렉토리 사용 + 추가 경로)
     *
     * @param reportData 리포트 데이터
     * @param paths      템플릿 파일 경로
     * @return 생성된 Workbook
     * @throws IOException 템플릿 파일을 읽는 중 오류 발생 시
     */
    public Workbook generateReport(ReportData reportData, String... paths) throws IOException {
        // 템플릿 파일 경로 설정
        Path templatePath = resolveTemplatePath(reportData.templateName(), paths);

        if (!Files.exists(templatePath)) {
            throw new IllegalArgumentException("Template file not found: " + reportData.templateName());
        }

        // 템플릿 로드
        Workbook workbook = ExcelFactory.loadTemplate(templatePath);

        // 전략 설정
        ExcelReportStrategy strategy = strategyMap.get(reportData.type());

        if (strategy == null) {
            throw new IllegalArgumentException("Unknown strategy type: " + reportData.type());
        }

        // 리포트 채우기
        strategy.fillReport(workbook, reportData);
        return workbook;
    }
}
