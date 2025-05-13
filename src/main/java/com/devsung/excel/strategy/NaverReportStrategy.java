package com.devsung.excel.strategy;

import com.devsung.excel.core.ExcelSheetUtils;
import com.devsung.excel.model.ReportData;
import com.devsung.excel.model.ReportSheetData;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import static com.devsung.excel.core.ExcelSheetUtils.isSheetEmpty;

public class NaverReportStrategy implements ExcelReportStrategy {

    @Override
    public void fillReport(Workbook workbook, ReportData reportData) {
        // 시트 데이터가 비어있는지 확인
        checkAndCleanSheetData(
                workbook,
                reportData.sheetData().stream().map(ReportSheetData::sheetName).toList()
        );

        reportData.sheetData().forEach(sheet -> {
            Sheet target = workbook.getSheet(sheet.sheetName());
            if (target == null) {
                target = workbook.createSheet(sheet.sheetName());
            }

            int rowIdx = 1;

            for (Object item : sheet.data()) {
                Row row = target.createRow(rowIdx++);
                Field[] fields = item.getClass().getDeclaredFields();
                int colIdx = 0;
                for (Field field : fields) {
                    field.setAccessible(true);
                    try {
                        Object value = field.get(item);
                        Cell cell = row.createCell(colIdx++);

                        if (value == null) {
                            cell.setCellValue("");
                        } else if (value instanceof Number) {
                            cell.setCellValue(((Number) value).doubleValue());
                        } else if (value instanceof java.time.LocalDate) {
                            String formattedDate = ((java.time.LocalDate) value).format(java.time.format.DateTimeFormatter.ofPattern("yyyy-MM-dd"));
                            cell.setCellValue(formattedDate);
                        } else {
                            cell.setCellValue(value.toString());
                        }

                    } catch (IllegalAccessException e) {
                        throw new RuntimeException("Failed to access field: " + field.getName(), e);
                    } finally {
                        field.setAccessible(false);
                    }
                }
            }
        });
    }

    /**
     * 시트 데이터가 비어있는지 확인, 데이터가 있으면 삭제
     *
     * @param workbook   엑셀 워크북
     * @param sheetNames 처리할 시트 이름 리스트
     */
    public void checkAndCleanSheetData(Workbook workbook, List<String> sheetNames) {
        List<String> sheetsToClean = new ArrayList<>();

        for (String sheetName : sheetNames) {
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet != null && !isSheetEmpty(sheet)) {
                sheetsToClean.add(sheetName);
            }
        }

        if (!sheetsToClean.isEmpty()) {
            ExcelSheetUtils.deleteSheetData(workbook, sheetsToClean);
        }
    }
}
