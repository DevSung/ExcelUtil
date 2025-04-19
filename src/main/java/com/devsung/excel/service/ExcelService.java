package com.devsung.excel.service;

import com.devsung.excel.core.ExcelProcessor;
import com.devsung.excel.core.PoiExcelProcessor;
import com.devsung.excel.model.SheetDataRequest;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.List;
import java.util.Map;

public class ExcelService {

    private final TemplateManager templateManager;
    private final ReportGenerator reportGenerator;
    private final ExcelProcessor processor;

    public ExcelService(File storageDir) {
        this.processor = new PoiExcelProcessor();
        this.templateManager = new TemplateManager(storageDir);
        this.reportGenerator = new ReportGenerator(processor);
    }

    public File saveTemplate(String fileName, InputStream inputStream) throws IOException {
        return templateManager.saveTemplate(fileName, inputStream);
    }

    public ByteArrayOutputStream generateReport(String templateName, String sheetName, List<Map<String, String>> data)
            throws IOException {
        File template = templateManager.getTemplate(templateName);
        return reportGenerator.generateReport(template, sheetName, data);
    }

    public ByteArrayOutputStream generateMultiSheetReport(String templateName, List<SheetDataRequest> sheets)
            throws IOException {

        File template = templateManager.getTemplate(templateName);

        try (InputStream in = new FileInputStream(template);
             Workbook workbook = processor.loadWorkbook(in)) {

            for (SheetDataRequest sheet : sheets) {
                processor.writeSheet(workbook, sheet.getSheetName(), sheet.getData());
            }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            processor.saveWorkbook(workbook, out);
            return out;
        }
    }

    public List<String> getSheetNames(String templateName) throws IOException {
        File template = templateManager.getTemplate(templateName);
        try (InputStream in = new FileInputStream(template);
             Workbook workbook = processor.loadWorkbook(in)) {
            return processor.getSheetNames(workbook);
        }
    }

    public List<Map<String, String>> readSheet(String templateName, String sheetName) throws IOException {
        File template = templateManager.getTemplate(templateName);
        try (InputStream in = new FileInputStream(template);
             Workbook workbook = processor.loadWorkbook(in)) {
            return processor.readSheet(workbook, sheetName);
        }
    }

    public List<String> listTemplates() {
        return templateManager.listTemplates();
    }

    public boolean deleteTemplate(String templateName) {
        return templateManager.deleteTemplate(templateName);
    }
}