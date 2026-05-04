package org.example.gov;


import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.gov.entity.SheetJson;
import org.example.gov.entity.WorkbookJson;
import org.example.gov.service.ExcelConversionException;
import org.example.gov.service.ExcelToJson;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

public class ExcelToJsonTest {

    private final ExcelToJson excelToJson = new ExcelToJson();

    @Test
    void convertShouldParseSimpleExcelFile() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("People");

            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Name");
            header.createCell(1).setCellValue("Age");
            header.createCell(2).setCellValue("Active");

            Row row = sheet.createRow(1);
            row.createCell(0).setCellValue("John");
            row.createCell(1).setCellValue(25);
            row.createCell(2).setCellValue(true);

            WorkbookJson result = excelToJson.convert(toMultipartFile(workbook, "people.xlsx"));

            assertEquals("people.xlsx", result.getFileName());
            assertEquals(1, result.getSheetCount());

            SheetJson resultSheet = result.getSheetJsonJsons().getFirst();

            assertEquals("People", resultSheet.getSheetName());
            assertEquals(0, resultSheet.getHeaderRowIndex());
            assertEquals(List.of("Name", "Age", "Active"), resultSheet.getColumns());
            assertEquals(1, resultSheet.getMapList().size());

            Map<String, Object> resultRow = resultSheet.getMapList().getFirst();

            assertEquals("John", resultRow.get("Name"));
            assertEquals(25.0, resultRow.get("Age"));
            assertEquals(true, resultRow.get("Active"));
        }
    }

    @Test
    void convertShouldTrimStringValues() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Data");

            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue(" Name ");

            Row row = sheet.createRow(1);
            row.createCell(0).setCellValue(" Alice ");

            WorkbookJson result = excelToJson.convert(toMultipartFile(workbook, "trim.xlsx"));

            SheetJson resultSheet = result.getSheetJsonJsons().getFirst();

            assertEquals(List.of("Name"), resultSheet.getColumns());
            assertEquals("Alice", resultSheet.getMapList().getFirst().get("Name"));
        }
    }

    @Test
    void convertShouldMakeDuplicateColumnNamesUnique() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Duplicates");

            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Name");
            header.createCell(1).setCellValue("Name");
            header.createCell(2).setCellValue("Name");

            Row row = sheet.createRow(1);
            row.createCell(0).setCellValue("A");
            row.createCell(1).setCellValue("B");
            row.createCell(2).setCellValue("C");

            WorkbookJson result = excelToJson.convert(toMultipartFile(workbook, "duplicates.xlsx"));

            SheetJson resultSheet = result.getSheetJsonJsons().getFirst();

            assertEquals(List.of("Name", "Name_2", "Name_3"), resultSheet.getColumns());

            Map<String, Object> resultRow = resultSheet.getMapList().getFirst();

            assertEquals("A", resultRow.get("Name"));
            assertEquals("B", resultRow.get("Name_2"));
            assertEquals("C", resultRow.get("Name_3"));
        }
    }

    @Test
    void convertShouldUseGeneratedColumnNameForBlankHeaderCell() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("GeneratedColumns");

            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Name");
            header.createCell(1).setCellValue("");

            Row row = sheet.createRow(1);
            row.createCell(0).setCellValue("John");
            row.createCell(1).setCellValue("Unknown value");

            WorkbookJson result = excelToJson.convert(toMultipartFile(workbook, "generated-columns.xlsx"));

            SheetJson resultSheet = result.getSheetJsonJsons().getFirst();

            assertEquals(List.of("Name", "column_2"), resultSheet.getColumns());
            assertEquals("Unknown value", resultSheet.getMapList().getFirst().get("column_2"));
        }
    }

    @Test
    void convertShouldDetectHeaderAfterEmptyRows() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("ShiftedHeader");

            sheet.createRow(0);
            sheet.createRow(1);

            Row header = sheet.createRow(2);
            header.createCell(0).setCellValue("Code");

            Row row = sheet.createRow(3);
            row.createCell(0).setCellValue("ABC-123");

            WorkbookJson result = excelToJson.convert(toMultipartFile(workbook, "shifted-header.xlsx"));

            SheetJson resultSheet = result.getSheetJsonJsons().getFirst();

            assertEquals(2, resultSheet.getHeaderRowIndex());
            assertEquals(List.of("Code"), resultSheet.getColumns());
            assertEquals("ABC-123", resultSheet.getMapList().getFirst().get("Code"));
        }
    }

    @Test
    void convertShouldReturnEmptySheetWhenNoHeaderExists() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("EmptySheet");

            WorkbookJson result = excelToJson.convert(toMultipartFile(workbook, "empty.xlsx"));

            SheetJson resultSheet = result.getSheetJsonJsons().getFirst();

            assertEquals("EmptySheet", resultSheet.getSheetName());
            assertEquals(-1, resultSheet.getHeaderRowIndex());
            assertTrue(resultSheet.getColumns().isEmpty());
            assertTrue(resultSheet.getMapList().isEmpty());
        }
    }

    @Test
    void convertShouldParseMultipleSheets() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            var firstSheet = workbook.createSheet("First");
            Row firstHeader = firstSheet.createRow(0);
            firstHeader.createCell(0).setCellValue("Name");
            Row firstRow = firstSheet.createRow(1);
            firstRow.createCell(0).setCellValue("John");

            var secondSheet = workbook.createSheet("Second");
            Row secondHeader = secondSheet.createRow(0);
            secondHeader.createCell(0).setCellValue("City");
            Row secondRow = secondSheet.createRow(1);
            secondRow.createCell(0).setCellValue("Kyiv");

            WorkbookJson result = excelToJson.convert(toMultipartFile(workbook, "multiple-sheets.xlsx"));

            assertEquals(2, result.getSheetCount());
            assertEquals("First", result.getSheetJsonJsons().get(0).getSheetName());
            assertEquals("Second", result.getSheetJsonJsons().get(1).getSheetName());

            assertEquals("John", result.getSheetJsonJsons().get(0).getMapList().getFirst().get("Name"));
            assertEquals("Kyiv", result.getSheetJsonJsons().get(1).getMapList().getFirst().get("City"));
        }
    }

    @Test
    void convertShouldParseDateCellsAsLocalDate() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Dates");

            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Date");

            CellStyle dateStyle = workbook.createCellStyle();
            CreationHelper creationHelper = workbook.getCreationHelper();
            dateStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-mm-dd"));

            Row row = sheet.createRow(1);
            var dateCell = row.createCell(0);
            dateCell.setCellValue(LocalDate.of(2026, 5, 4));
            dateCell.setCellStyle(dateStyle);

            WorkbookJson result = excelToJson.convert(toMultipartFile(workbook, "dates.xlsx"));

            Object value = result.getSheetJsonJsons()
                    .getFirst()
                    .getMapList()
                    .getFirst()
                    .get("Date");

            assertEquals(LocalDate.of(2026, 5, 4), value);
        }
    }

    @Test
    void convertShouldParseFormulaCellUsingCachedResult() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Formulas");

            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Total");

            Row row = sheet.createRow(1);
            row.createCell(0).setCellFormula("2+3");

            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();

            WorkbookJson result = excelToJson.convert(toMultipartFile(workbook, "formulas.xlsx"));

            Object value = result.getSheetJsonJsons()
                    .getFirst()
                    .getMapList()
                    .getFirst()
                    .get("Total");

            assertEquals(5.0, value);
        }
    }

    @Test
    void convertShouldRejectEmptyFile() {
        MultipartFile file = new MockMultipartFile(
                "file",
                "empty.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                new byte[0]
        );

        RuntimeException exception = assertThrows(RuntimeException.class, () -> excelToJson.convert(file));

        assertEquals("Файл порожній", exception.getMessage());
    }

    @Test
    void convertShouldRejectUnsupportedExtension() {
        MultipartFile file = new MockMultipartFile(
                "file",
                "data.txt",
                "text/plain",
                "not excel".getBytes()
        );

        RuntimeException exception = assertThrows(RuntimeException.class, () -> excelToJson.convert(file));

        assertEquals("Підтримуються тільки Excel файли", exception.getMessage());
    }

    @Test
    void convertShouldWrapInvalidExcelContentIntoExcelConversionException() {
        MultipartFile file = new MockMultipartFile(
                "file",
                "broken.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "not a real excel file".getBytes()
        );

        ExcelConversionException exception = assertThrows(
                ExcelConversionException.class,
                () -> excelToJson.convert(file)
        );

        assertEquals("Файл має некоректний або непідтримуваний Excel формат.", exception.getMessage());
        assertNotNull(exception.getCause());
    }

    private MultipartFile toMultipartFile(Workbook workbook, String fileName) throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);

        return new MockMultipartFile(
                "file",
                fileName,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                outputStream.toByteArray()
        );
    }
}

