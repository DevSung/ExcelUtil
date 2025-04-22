package com.devsung.excel.model;

import org.apache.poi.ss.formula.functions.T;

import java.util.List;

public record ReportSheetData(
        String sheetName,
        List<T> data
) {
    public ReportSheetData {
        if (sheetName == null || sheetName.isEmpty()) {
            throw new IllegalArgumentException("Sheet name cannot be null or empty");
        }
    }
}
