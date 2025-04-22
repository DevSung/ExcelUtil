package com.devsung.excel.core;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Set;

public class ExcelFileUtils {

    private static final Set<String> ALLOWED_EXTENSIONS = Set.of(".xlsx", ".xls");
    private static final String SPECIAL_CHARS_PATTERN = "[\\\\/:*?\"<>|]";
    private static final String REPLACEMENT_CHAR = "_";

    /**
     * 엑셀 템플릿 파일을 저장합니다.
     *
     * @param templateName 저장할 템플릿 파일 이름 (.xlsx 또는 .xls만 허용)
     * @param inputStream  템플릿 파일을 읽어올 InputStream
     * @param savePath     템플릿 파일을 저장할 경로
     * @throws IOException 파일 복사 중 오류가 발생했을 경우
     */
    public static void save(String templateName, InputStream inputStream, Path savePath) throws IOException {
        validateTemplateName(templateName);

        // 파일명에서 특수문자를 치환
        String sanitizedTemplateName = sanitizeFileName(templateName);

        // 새로운 저장 경로 생성
        Path sanitizedSavePath = savePath.resolveSibling(sanitizedTemplateName);

        // 디렉토리가 존재하지 않으면 생성
        if (!Files.exists(sanitizedSavePath.getParent())) {
            Files.createDirectories(sanitizedSavePath.getParent());
        }

        // 파일 복사
        Files.copy(inputStream, sanitizedSavePath, StandardCopyOption.REPLACE_EXISTING);
    }

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
     * 엑셀 템플릿 파일을 삭제합니다.
     *
     * @param path 삭제할 템플릿 파일 경로
     * @return 삭제 성공 여부
     * @throws IOException 파일 삭제 중 오류가 발생했을 경우
     */
    public static boolean delete(Path path) throws IOException {
        return Files.exists(path) && Files.deleteIfExists(path);
    }

    /**
     * 템플릿 파일이 존재하는지 확인
     *
     * @param templatePath 확인할 템플릿 경로
     * @return 존재 여부
     */
    public static boolean exists(Path templatePath) {
        return Files.exists(templatePath);
    }

    /**
     * 템플릿 이름이 유효한지 검증합니다.
     *
     * @param templateName 검증할 템플릿 이름
     */
    private static void validateTemplateName(String templateName) {
        if (templateName == null) {
            throw new IllegalArgumentException("Template name cannot be null");
        }

        boolean validExtension = ALLOWED_EXTENSIONS.stream()
                .anyMatch(templateName::endsWith);

        if (!validExtension) {
            throw new IllegalArgumentException("Only " + ALLOWED_EXTENSIONS + " files are allowed");
        }
    }

    /**
     * 파일명에 포함된 특수문자를 치환합니다.
     *
     * @param fileName 원본 파일명
     * @return 치환된 파일명
     */
    private static String sanitizeFileName(String fileName) {
        // 특수문자 \ / : * ? " < > | 를 _ 로 치환
        return fileName.replaceAll(SPECIAL_CHARS_PATTERN, REPLACEMENT_CHAR);
    }

}
