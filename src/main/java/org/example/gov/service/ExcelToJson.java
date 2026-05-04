package org.example.gov.service;

import org.apache.poi.ss.usermodel.*;
import org.example.gov.entity.SheetJson;
import org.example.gov.entity.WorkbookJson;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

@Service
public class ExcelToJson {

    private static final int MAX_EMPTY_ROWS_AFTER_HEADER = 50;

    public WorkbookJson convert(MultipartFile file) {
        validate(file);

        try (InputStream is = file.getInputStream();
             Workbook workbook = WorkbookFactory.create(is)) {

            List<SheetJson> sheets = new ArrayList<>();

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheets.add(processSheet(workbook.getSheetAt(i)));
            }

            return new WorkbookJson(
                    file.getOriginalFilename(),
                    sheets.size(),
                    sheets
            );

        } catch (IOException e) {
            throw new ExcelConversionException("Не вдалося прочитати Excel файл.", e);
        } catch (Exception e) {
            throw new ExcelConversionException("Файл має некоректний або непідтримуваний Excel формат.", e);
        }
    }

    private SheetJson processSheet(Sheet sheet) {

        int headerRowIndex = findHeaderRow(sheet);
        if (headerRowIndex == -1) {
            return new SheetJson(sheet.getSheetName(), -1, List.of(), List.of());
        }

        Row headerRow = sheet.getRow(headerRowIndex);
        int columnCount = headerRow.getLastCellNum();

        List<String> columns = buildColumns(headerRow, columnCount);
        List<Map<String, Object>> data = new ArrayList<>();

        int emptyCount = 0;

        for (int i = headerRowIndex + 1; i <= sheet.getLastRowNum(); i++) {

            Row row = sheet.getRow(i);

            if (isRowEmpty(row, columnCount)) {
                emptyCount++;
                if (emptyCount >= MAX_EMPTY_ROWS_AFTER_HEADER) break;
                continue;
            }

            emptyCount = 0;

            Map<String, Object> rowMap = new LinkedHashMap<>();

            for (int j = 0; j < columnCount; j++) {
                rowMap.put(columns.get(j), getCellValue(row, j));
            }

            data.add(rowMap);
        }

        return new SheetJson(
                sheet.getSheetName(),
                headerRowIndex,
                columns,
                data
        );
    }

    private int findHeaderRow(Sheet sheet) {
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (!isRowEmpty(row, 10)) { // перевіряємо тільки перші 10 клітинок
                return i;
            }
        }
        return -1;
    }

    private List<String> buildColumns(Row headerRow, int columnCount) {
        List<String> columns = new ArrayList<>();

        for (int i = 0; i < columnCount; i++) {
            String name = Optional.ofNullable(getCellValue(headerRow, i))
                    .map(Object::toString)
                    .map(String::trim)
                    .filter(s -> !s.isBlank())
                    .orElse("column_" + (i + 1));

            columns.add(makeUnique(name, columns));
        }

        return columns;
    }

    private String makeUnique(String base, List<String> existing) {
        String result = base;
        int i = 2;

        while (existing.contains(result)) {
            result = base + "_" + i++;
        }

        return result;
    }

    private boolean isRowEmpty(Row row, int columnCount) {
        if (row == null) return true;

        for (int i = 0; i < columnCount; i++) {
            Cell cell = row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                return false;
            }
        }
        return true;
    }

    private Object getCellValue(Row row, int columnIndex) {
        if (row == null) return null;

        Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return null;

        return switch (cell.getCellType()) {

            case STRING -> cell.getStringCellValue().trim();

            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    yield cell.getLocalDateTimeCellValue().toLocalDate();
                }
                yield cell.getNumericCellValue();
            }

            case BOOLEAN -> cell.getBooleanCellValue();

            case FORMULA -> evaluateFormula(cell);

            default -> null;
        };
    }

    private Object evaluateFormula(Cell cell) {
        try {
            return switch (cell.getCachedFormulaResultType()) {
                case STRING -> cell.getStringCellValue();
                case NUMERIC -> cell.getNumericCellValue();
                case BOOLEAN -> cell.getBooleanCellValue();
                default -> null;
            };
        } catch (Exception e) {
            return null;
        }
    }

    private void validate(MultipartFile file) {
        if (file == null || file.isEmpty()) {
            throw new RuntimeException("Файл порожній");
        }

        String name = Optional.ofNullable(file.getOriginalFilename())
                .orElse("")
                .toLowerCase();

        if (!name.endsWith(".xls") && !name.endsWith(".xlsx")) {
            throw new RuntimeException("Підтримуються тільки Excel файли");
        }
    }
}