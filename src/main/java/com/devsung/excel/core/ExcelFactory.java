package com.devsung.excel.core;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

public class ExcelFactory {

    /**
     * 주어진 경로에서 템플릿 파일을 로드합니다.
     *
     * @param path 엑셀 템플릿 파일의 경로
     * @return 로드된 Workbook 객체
     * @throws IOException 템플릿 파일을 읽는 중 오류가 발생한 경우
     */
    public static Workbook loadTemplate(Path path) throws IOException {
        try (InputStream in = Files.newInputStream(path)) {
            return WorkbookFactory.create(in);
        }
    }

}
