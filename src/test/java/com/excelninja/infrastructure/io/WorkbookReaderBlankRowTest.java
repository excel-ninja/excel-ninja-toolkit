package com.excelninja.infrastructure.io;

import com.excelninja.domain.annotation.ExcelReadColumn;
import com.excelninja.domain.model.ExcelSheet;
import com.excelninja.domain.model.ExcelWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

@DisplayName("Workbook readers skip blank rows consistently")
class WorkbookReaderBlankRowTest {

    @TempDir
    Path tempDir;

    static class EmployeeRow {
        @ExcelReadColumn(headerName = "ID")
        private Long id;

        @ExcelReadColumn(headerName = "Name")
        private String name;

        public Long getId() {
            return id;
        }

        public String getName() {
            return name;
        }
    }

    @Test
    @DisplayName("POI reader skips blank, whitespace and empty-formula rows")
    void poiReaderSkipsBlankRows() throws Exception {
        File workbookFile = createWorkbookWithBlankRows("poi_blank_rows.xlsx");

        ExcelWorkbook workbook = new PoiWorkbookReader().read(workbookFile);
        ExcelSheet sheet = workbook.getSheet("Employees");

        assertThat(sheet.getRows().size()).isEqualTo(2);
        assertThat(sheet.getCellValue(0, "Name")).isEqualTo("Alice");
        assertThat(sheet.getCellValue(1, "Name")).isEqualTo("Bob");
    }

    @Test
    @DisplayName("Streaming reader skips blank, whitespace and empty-formula rows")
    void streamingReaderSkipsBlankRows() throws Exception {
        File workbookFile = createWorkbookWithBlankRows("streaming_blank_rows.xlsx");

        ExcelWorkbook workbook = new StreamingWorkbookReader().read(workbookFile);
        ExcelSheet sheet = workbook.getSheet("Employees");

        assertThat(sheet.getRows().size()).isEqualTo(2);
        assertThat(sheet.getCellValue(0, "Name")).isEqualTo("Alice");
        assertThat(sheet.getCellValue(1, "Name")).isEqualTo("Bob");
    }

    @Test
    @DisplayName("Chunk reader does not emit empty entities for blank rows")
    void chunkReaderSkipsBlankRows() throws Exception {
        File workbookFile = createWorkbookWithBlankRows("chunk_blank_rows.xlsx");
        Iterator<List<EmployeeRow>> chunks = new StreamingWorkbookReader().readInChunks(workbookFile, EmployeeRow.class, 1);
        List<EmployeeRow> rows = new ArrayList<>();

        try {
            while (chunks.hasNext()) {
                rows.addAll(chunks.next());
            }
        } finally {
            ((AutoCloseable) chunks).close();
        }

        assertThat(rows).hasSize(2);
        assertThat(rows).extracting(EmployeeRow::getId).containsExactly(1L, 2L);
        assertThat(rows).extracting(EmployeeRow::getName).containsExactly("Alice", "Bob");
    }

    private File createWorkbookWithBlankRows(String fileName) throws IOException {
        Path workbookPath = tempDir.resolve(fileName);

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Employees");

            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("ID");
            headerRow.createCell(1).setCellValue("Name");

            Row aliceRow = sheet.createRow(1);
            aliceRow.createCell(0).setCellValue(1);
            aliceRow.createCell(1).setCellValue("Alice");

            sheet.createRow(2);

            Row whitespaceRow = sheet.createRow(3);
            whitespaceRow.createCell(0).setCellValue("   ");
            whitespaceRow.createCell(1).setCellValue(" ");

            Row emptyFormulaRow = sheet.createRow(4);
            emptyFormulaRow.createCell(0).setCellFormula("\"\"");
            emptyFormulaRow.createCell(1).setCellFormula("\"\"");

            Row bobRow = sheet.createRow(5);
            bobRow.createCell(0).setCellValue(2);
            bobRow.createCell(1).setCellValue("Bob");

            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();

            try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
                workbook.write(outputStream);
            }
        }

        return workbookPath.toFile();
    }
}
