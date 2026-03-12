package com.excelninja.infrastructure.io;

import com.excelninja.domain.exception.InvalidDocumentStructureException;
import com.excelninja.domain.model.ExcelWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

@DisplayName("POI reader header conversion")
class PoiWorkbookReaderHeaderTest {

    @TempDir
    Path tempDir;

    @Test
    @DisplayName("Numeric, boolean and formula headers are converted to strings")
    void shouldConvertMixedHeaderTypesToStrings() throws Exception {
        Path workbookPath = tempDir.resolve("mixed_headers.xlsx");

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("MixedHeaders");
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue(101);
            headerRow.createCell(1).setCellValue("Name");
            headerRow.createCell(2).setCellValue(true);
            headerRow.createCell(3).setCellFormula("1+1");

            Row dataRow = sheet.createRow(1);
            dataRow.createCell(0).setCellValue("A");
            dataRow.createCell(1).setCellValue("Alice");
            dataRow.createCell(2).setCellValue("Y");
            dataRow.createCell(3).setCellValue("Ready");

            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();

            try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
                workbook.write(outputStream);
            }
        }

        ExcelWorkbook workbook = new PoiWorkbookReader().read(workbookPath.toFile());

        assertThat(workbook.getSheet("MixedHeaders").getHeaders().getHeaderNames())
                .containsExactly("101", "Name", "true", "2");
    }

    @Test
    @DisplayName("Headers that become blank after conversion are rejected")
    void shouldRejectBlankHeadersAfterConversion() throws IOException {
        Path workbookPath = tempDir.resolve("blank_formula_header.xlsx");

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("InvalidHeaders");
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellFormula("\"\"");
            headerRow.createCell(1).setCellValue("Name");
            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();

            try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
                workbook.write(outputStream);
            }
        }

        assertThatThrownBy(() -> new PoiWorkbookReader().read(workbookPath.toFile()))
                .isInstanceOf(InvalidDocumentStructureException.class)
                .hasMessageContaining("Header cannot be empty");
    }
}
