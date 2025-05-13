package com.devsung.excel.model;

import java.util.List;

public record ReportSheetData<T>(
        String sheetName,
        List<T> data
) {
    public ReportSheetData {
        if (sheetName == null || sheetName.isEmpty()) {
            throw new IllegalArgumentException("Sheet name cannot be null or empty");
        }
    }
}
