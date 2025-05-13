package com.devsung.excel.model;

import org.apache.commons.collections4.CollectionUtils;

import java.util.List;

public record ReportData(
        ReportType type,
        String templateName,
        List<ReportSheetData<?>> sheetData
) {
    public ReportData {
        if (templateName == null || templateName.isEmpty()) {
            throw new IllegalArgumentException("Template name cannot be null or empty");
        }
        if (CollectionUtils.isEmpty(sheetData)) {
            throw new IllegalArgumentException("Sheet data cannot be null or empty");
        }
    }
}
