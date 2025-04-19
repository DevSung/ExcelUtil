package com.devsung.excel.core;

import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;

public class PoiExcelProcessor extends AbstractExcelProcessor {

    @Override
    public Workbook loadWorkbook(InputStream inputStream) throws IOException {
        return WorkbookFactory.create(inputStream);
    }

    @Override
    public void saveWorkbook(Workbook workbook, OutputStream outputStream) throws IOException {
        workbook.write(outputStream);
        outputStream.flush();
    }

    @Override
    public List<Map<String, String>> readSheet(Workbook workbook, String sheetName) {
        List<Map<String, String>> data = new ArrayList<>();
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) return data;

        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return data;

        List<String> headers = new ArrayList<>();
        for (Cell cell : headerRow) {
            headers.add(cell.getStringCellValue());
        }

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;
            Map<String, String> rowData = new HashMap<>();
            for (int j = 0; j < headers.size(); j++) {
                Cell cell = row.getCell(j);
                rowData.put(headers.get(j), cell != null ? cell.toString() : "");
            }
            data.add(rowData);
        }

        return data;
    }

    @Override
    public void writeSheet(Workbook workbook, String sheetName, List<Map<String, String>> data) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            sheet = workbook.createSheet(sheetName);
        }

        if (data == null || data.isEmpty()) return;

        Set<String> headers = data.get(0).keySet();
        int rowIdx = 0;

        // 헤더 작성
        Row headerRow = sheet.createRow(rowIdx++);
        int colIdx = 0;
        Map<String, Integer> headerIndexMap = new LinkedHashMap<>();
        for (String header : headers) {
            Cell cell = headerRow.createCell(colIdx);
            cell.setCellValue(header);
            headerIndexMap.put(header, colIdx);
            colIdx++;
        }

        // 데이터 작성
        for (Map<String, String> rowData : data) {
            Row row = sheet.createRow(rowIdx++);
            for (Map.Entry<String, String> entry : rowData.entrySet()) {
                int col = headerIndexMap.get(entry.getKey());
                Cell cell = row.createCell(col);
                cell.setCellValue(entry.getValue());
            }
        }

        for (int i = 0; i < headerIndexMap.size(); i++) {
            sheet.autoSizeColumn(i);
        }
    }

}
