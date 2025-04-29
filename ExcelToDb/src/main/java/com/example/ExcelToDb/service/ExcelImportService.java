package com.example.ExcelToDb.service;

import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public interface ExcelImportService {
    void importExcelToDatabase(MultipartFile file) throws IOException;

    List<Map<String, String>> readExcelFile(MultipartFile file) throws IOException;
}
