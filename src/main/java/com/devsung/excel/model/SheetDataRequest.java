package com.devsung.excel.model;

import java.util.List;
import java.util.Map;

public class SheetDataRequest {
    private String sheetName;
    private List<Map<String, String>> data;

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public List<Map<String, String>> getData() {
        return data;
    }

    public void setData(List<Map<String, String>> data) {
        this.data = data;
    }
}
