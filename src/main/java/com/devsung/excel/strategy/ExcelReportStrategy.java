package com.devsung.excel.strategy;

import com.devsung.excel.model.ReportData;
import org.apache.poi.ss.usermodel.Workbook;

public interface ExcelReportStrategy {
    void fillReport(Workbook workbook, ReportData reportData);
}
