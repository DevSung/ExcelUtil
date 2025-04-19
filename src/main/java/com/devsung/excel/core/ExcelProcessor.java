package com.devsung.excel.core;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

public interface ExcelProcessor {
    Workbook loadWorkbook(InputStream inputStream) throws IOException;

    void saveWorkbook(Workbook workbook, OutputStream outputStream) throws IOException;

    List<String> getSheetNames(Workbook workbook);

    String getFileName(File file);

    List<Map<String, String>> readSheet(Workbook workbook, String sheetName);

    void writeSheet(Workbook workbook, String sheetName, List<Map<String, String>> data);
}
