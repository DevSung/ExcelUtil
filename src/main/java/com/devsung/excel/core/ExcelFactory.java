package com.devsung.excel.core;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

public class ExcelFactory {


    /**
     * 주어진 Workbook 객체를 출력 스트림에 작성합니다.
     *
     * @param workbook 작성할 Workbook 객체
     * @param out      출력 스트림
     * @throws IOException 워크북 작성 중 오류가 발생한 경우
     */
    public static void write(Workbook workbook, OutputStream out) throws IOException {
        try (workbook) {
            workbook.write(out);
        }
    }

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

    /**
     * InputStream을 Workbook 객체로 변환합니다.
     *
     * @param inputStream 엑셀 파일 스트림
     * @return 변환된 Workbook 객체
     * @throws IOException 파일 읽기 중 오류 발생 시
     */
    public static Workbook convertFromStream(InputStream inputStream) throws IOException {
        if (inputStream == null) {
            throw new IllegalArgumentException("InputStream cannot be null");
        }
        return WorkbookFactory.create(inputStream);
    }

}
