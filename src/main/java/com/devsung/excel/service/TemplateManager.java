package com.devsung.excel.service;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class TemplateManager {

    private final File storageDir;

    public TemplateManager(File storageDir) {
        this.storageDir = storageDir;
        if (!this.storageDir.exists()) {
            this.storageDir.mkdirs();
        }
    }

    public File saveTemplate(String fileName, InputStream inputStream) throws IOException {
        File targetFile = new File(storageDir, fileName);
        try (OutputStream out = new FileOutputStream(targetFile)) {
            byte[] buffer = new byte[8192];
            int bytesRead;
            while ((bytesRead = inputStream.read(buffer)) != -1) {
                out.write(buffer, 0, bytesRead);
            }
        }
        return targetFile;
    }

    public File getTemplate(String fileName) {
        File file = new File(storageDir, fileName);
        if (!file.exists()) {
            throw new IllegalArgumentException("Template file not found: " + fileName);
        }
        return file;
    }

    public List<String> listTemplates() {
        File[] files = storageDir.listFiles();
        List<String> list = new ArrayList<>();
        if (files != null) {
            for (File file : files) {
                if (file.isFile()) {
                    list.add(file.getName());
                }
            }
        }
        return list;
    }

    public boolean deleteTemplate(String fileName) {
        File file = new File(storageDir, fileName);
        return file.exists() && file.delete();
    }

}
