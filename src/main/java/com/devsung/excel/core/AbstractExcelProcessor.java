package com.devsung.excel.core;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public abstract class AbstractExcelProcessor implements ExcelProcessor {

    @Override
    public String getFileName(File file) {
        return file.getName();
    }

    @Override
    public List<String> getSheetNames(Workbook workbook) {
        List<String> names = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            names.add(workbook.getSheetName(i));
        }
        return names;
    }

}
