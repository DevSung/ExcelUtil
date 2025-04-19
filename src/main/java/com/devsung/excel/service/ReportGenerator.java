package com.devsung.excel.service;

import com.devsung.excel.core.ExcelProcessor;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.List;
import java.util.Map;

public class ReportGenerator {
    private final ExcelProcessor processor;

    public ReportGenerator(ExcelProcessor processor) {
        this.processor = processor;
    }

    public ByteArrayOutputStream generateReport(File templateFile, String sheetName, List<Map<String, String>> data)
            throws IOException {

        try (InputStream in = new FileInputStream(templateFile);
             Workbook workbook = processor.loadWorkbook(in)) {

            processor.writeSheet(workbook, sheetName, data);

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            processor.saveWorkbook(workbook, out);
            return out;
        }
    }
}
