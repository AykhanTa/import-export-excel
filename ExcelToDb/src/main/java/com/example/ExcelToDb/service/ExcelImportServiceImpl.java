package com.example.ExcelToDb.service;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

@Service
@RequiredArgsConstructor
public class ExcelImportServiceImpl implements ExcelImportService {

    private final JdbcTemplate jdbcTemplate;

    @Transactional
    public void importExcelToDatabase(MultipartFile file) throws IOException {
        if (!Objects.requireNonNull(file.getOriginalFilename()).endsWith(".xlsx"))
            throw new IllegalArgumentException("Yalnız .xlsx fayllar qəbul olunur.");

        List<Map<String, String>> data = readExcelFile(file);
        if (data.isEmpty()) return;
        String tableName = "dynamic_table";

        jdbcTemplate.execute("DROP TABLE IF EXISTS " + tableName);
        jdbcTemplate.execute(generateCreateTableSql(tableName, data.get(0)));
        jdbcTemplate.batchUpdate(generateInsertSql(tableName, data.get(0)), prepareBatchArgs(data));
    }

    public List<Map<String, String>> readExcelFile(MultipartFile file) throws IOException {
        List<Map<String, String>> data = new ArrayList<>();
        try (InputStream inputStream = file.getInputStream()) {
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            List<String> headers = extractHeaders(sheet.getRow(0));

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                Map<String, String> rowData = new HashMap<>();
                for (int j = 0; j < headers.size(); j++) {
                    Cell cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    rowData.put(headers.get(j), cell.toString().trim());
                }
                data.add(rowData);
            }
        }
        return data;
    }

    private List<String> extractHeaders(Row headerRow) {
        List<String> headers = new ArrayList<>();
        for (Cell cell : headerRow) {
            headers.add(cell.getStringCellValue().trim());
        }
        return headers;
    }

    private String generateCreateTableSql(String tableName, Map<String, String> headers) {
        StringBuilder sql = new StringBuilder("CREATE TABLE ").append(tableName).append(" (");
        headers.keySet().forEach(header -> sql.append("\"").append(header).append("\" TEXT, "));
        sql.delete(sql.length() - 2, sql.length()).append(")");
        return sql.toString();
    }

    private String generateInsertSql(String tableName, Map<String, String> headers) {
        StringBuilder columnNames = new StringBuilder();
        StringBuilder values = new StringBuilder();
        headers.keySet().forEach(header -> {
            columnNames.append("\"").append(header).append("\", ");
            values.append("?, ");
        });
        columnNames.delete(columnNames.length() - 2, columnNames.length());
        values.delete(values.length() - 2, values.length());
        return "INSERT INTO " + tableName + " (" + columnNames + ") VALUES (" + values + ")";
    }

    private List<Object[]> prepareBatchArgs(List<Map<String, String>> data) {
        List<Object[]> batchArgs = new ArrayList<>();
        for (Map<String, String> row : data) {
            batchArgs.add(row.values().toArray());
        }
        return batchArgs;
    }
}
