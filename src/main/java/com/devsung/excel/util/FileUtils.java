package com.devsung.excel.util;

import java.io.File;

public class FileUtils {
    public static void ensureDirectoryExists(File dir) {
        if (!dir.exists()) {
            dir.mkdirs();
        }
    }
}