package com.example.ExcelToDb.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

@Service
public class ExcelExportServiceImpl implements ExcelExportService {

    private final JdbcTemplate jdbcTemplate;

    private final String exportDirectory = "excel-files";

    public ExcelExportServiceImpl(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }
    public void exportTableToDownloads() throws IOException {
        List<Map<String, Object>> rows = fetchDataFromDatabase();
        if (rows.isEmpty()) {
            throw new IOException("Table is empty! No data to export.");
        }

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Data");

        Map<String, Object> firstRow = rows.get(0);
        Row headerRow = sheet.createRow(0);
        int headerIndex = 0;
        for (String column : firstRow.keySet()) {
            headerRow.createCell(headerIndex++).setCellValue(column);
        }

        int rowIndex = 1;
        for (Map<String, Object> row : rows) {
            Row excelRow = sheet.createRow(rowIndex++);
            int columnIndex = 0;
            for (Object value : row.values()) {
                excelRow.createCell(columnIndex++).setCellValue(value != null ? value.toString() : "");
            }
        }
        String userHome = System.getProperty("user.home");
        String downloadsPath = userHome + File.separator + "Downloads";
        String fileName = "exported_data_" + System.currentTimeMillis() + ".xlsx";
        File file = new File(downloadsPath, fileName);

        try (FileOutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        }

        workbook.close();
    }

    private List<Map<String, Object>> fetchDataFromDatabase() {
        String sql = "SELECT * FROM dynamic_table";
        return jdbcTemplate.queryForList(sql);
    }
}
