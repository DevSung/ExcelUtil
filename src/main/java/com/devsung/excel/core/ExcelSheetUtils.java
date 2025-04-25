package com.devsung.excel.core;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;

public class ExcelSheetUtils {

    /**
     * 시트 이름을 기준으로 시트를 가져옵니다.
     *
     * @param workbook  엑셀 워크북
     * @param sheetName 시트 이름
     * @return 해당 시트, 없으면 null
     */
    public static Sheet getSheetByName(Workbook workbook, String sheetName) {
        return workbook.getSheet(sheetName);
    }

    /**
     * 워크북 내 모든 시트의 이름을 반환합니다.
     *
     * @param workbook 엑셀 워크북
     * @return 시트 이름 목록
     */
    public static List<String> getSheetNames(Workbook workbook) {
        List<String> names = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            names.add(workbook.getSheetName(i));
        }
        return names;
    }

    /**
     * 시트가 비어 있는지 확인합니다.
     *
     * @param sheet 엑셀 시트
     * @return 비어 있으면 true, 아니면 false
     */
    public static boolean isSheetEmpty(Sheet sheet) {
        return sheet.getPhysicalNumberOfRows() == 0;
    }

    /**
     * 시트의 1행(헤더)을 제외하고 데이터가 있는지 확인합니다.
     *
     * @param sheet 엑셀 시트
     * @return 1행을 제외한 데이터가 하나도 없으면 true(비어있음), 있으면 false
     */
    public static boolean isSheetBodyEmpty(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        // 데이터가 헤더만 있는 경우 (행이 1개 이하)
        if (lastRowNum < 1) {
            return true;
        }
        // 2번째 행부터 마지막 행까지 데이터가 실제로 있는지 확인
        for (int i = 1; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row != null && row.getPhysicalNumberOfCells() > 0) {
                return false;
            }
        }
        return true;
    }

    /**
     * 시트의 모든 데이터를 삭제합니다. (1행 제외 - 헤더)
     *
     * @param workbook  엑셀 워크북
     * @param sheetName 시트 이름
     * @return 삭제 성공 여부
     */
    public static boolean deleteSheetData(Workbook workbook, String sheetName) {
        Sheet sheet = getSheetByName(workbook, sheetName);
        if (sheet == null) {
            return false;
        }

        int lastRowNum = sheet.getLastRowNum();
        for (int i = lastRowNum; i > 0; i--) {
            sheet.removeRow(sheet.getRow(i));
        }
        return true;
    }

}
