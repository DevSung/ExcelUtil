package com.devsung.excel.model;

import java.util.List;
import java.util.Map;

public record SheetData(
        String sheetName,
        List<Map<String, String>> data
) {

}
