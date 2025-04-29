package com.example.ExcelToDb.controller;

import com.example.ExcelToDb.service.ExcelExportService;
import com.example.ExcelToDb.service.ExcelImportService;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController
@RequestMapping("/api/excel")
public class DataController {
    private final ExcelImportService excelImportService;

    private final ExcelExportService excelExportService;

    public DataController(ExcelImportService excelImportService, ExcelExportService excelExportService) {
        this.excelImportService = excelImportService;
        this.excelExportService = excelExportService;

    }

    @PostMapping("/import")
    public ResponseEntity<String> importExcel(@RequestParam("file") MultipartFile file) {
        try {
            excelImportService.importExcelToDatabase(file);
            return ResponseEntity.ok("Excel data imported successfully!");
        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Error reading Excel file");
        }
    }

    @GetMapping("/export")
    public String exportToExcel() {
        try {
            excelExportService.exportTableToDownloads();
            return "Excel faylı uğurla Downloads qovluğuna yazıldı.";
        } catch (Exception e) {
            return "Xəta baş verdi: " + e.getMessage();
        }
    }
}
