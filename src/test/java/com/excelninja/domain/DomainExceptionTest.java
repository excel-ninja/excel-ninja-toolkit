package com.excelninja.domain;

import com.excelninja.application.facade.NinjaExcel;
import com.excelninja.domain.annotation.ExcelReadColumn;
import com.excelninja.domain.annotation.ExcelWriteColumn;
import com.excelninja.domain.exception.DocumentConversionException;
import com.excelninja.domain.exception.EntityMappingException;
import com.excelninja.domain.exception.HeaderMismatchException;
import com.excelninja.domain.exception.InvalidDocumentStructureException;
import com.excelninja.domain.model.ExcelDocument;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.AssertionsForClassTypes.assertThatThrownBy;

class DomainExceptionTest {

    @TempDir
    Path tempDir;

    @Test
    @DisplayName("존재하지 않는 파일 읽기 시 DocumentConversionException 발생")
    void readNonExistentFile() {
        File nonExistentFile = new File("non_existent_file.xlsx");

        assertThatThrownBy(() -> NinjaExcel.read(nonExistentFile, ValidReadDto.class))
                .isInstanceOf(DocumentConversionException.class)
                .hasMessageContaining("does not exist");
    }

    @Test
    @DisplayName("null 파일 파라미터 시 DocumentConversionException 발생")
    void readNullFile() {
        assertThatThrownBy(() -> NinjaExcel.read((File) null, ValidReadDto.class))
                .isInstanceOf(DocumentConversionException.class)
                .hasMessageContaining("cannot be null");
    }

    @Test
    @DisplayName("null 스트링 파라미터 시 DocumentConversionException 발생")
    void readNullString() {
        assertThatThrownBy(() -> NinjaExcel.read((String) null, ValidReadDto.class))
                .isInstanceOf(DocumentConversionException.class)
                .hasMessageContaining("cannot be null");
    }

    @Test
    @DisplayName("어노테이션이 없는 클래스 사용 시 EntityMappingException 발생")
    void readWithoutAnnotation() throws IOException {
        File validExcel = createValidExcelFile();

        assertThatThrownBy(() -> NinjaExcel.read(validExcel, NoAnnotationDto.class))
                .isInstanceOf(EntityMappingException.class)
                .hasMessageContaining("No @ExcelReadColumn");
    }

    @Test
    @DisplayName("중복 헤더명 어노테이션 시 EntityMappingException 발생")
    void duplicateHeaderAnnotation() {
        List<DuplicateHeaderDto> data = Arrays.asList(new DuplicateHeaderDto("test1", "test2"));

        assertThatThrownBy(() -> ExcelDocument.writer().objects(data).create())
                .isInstanceOf(EntityMappingException.class)
                .hasMessageContaining("Duplicate header name");
    }

    @Test
    @DisplayName("빈 헤더명 어노테이션 시 EntityMappingException 발생")
    void emptyHeaderAnnotation() {
        List<EmptyHeaderDto> data = Arrays.asList(
                new EmptyHeaderDto("test")
        );

        assertThatThrownBy(() -> ExcelDocument.writer().objects(data).create())
                .isInstanceOf(EntityMappingException.class)
                .hasMessageContaining("Empty header name");
    }

    @Test
    @DisplayName("빈 엔티티 리스트로 문서 생성 시 EntityMappingException 발생")
    void emptyEntityList() {
        List<ValidWriteDto> emptyList = Arrays.asList();

        assertThatThrownBy(() -> ExcelDocument.writer().objects(emptyList).create())
                .isInstanceOf(EntityMappingException.class)
                .hasMessageContaining("cannot be null or empty");
    }

    @Test
    @DisplayName("ExcelDocument 빌더에서 빈 헤더 리스트 시 InvalidDocumentStructureException 발생")
    void emptyHeadersInBuilder() {
        assertThatThrownBy(() ->
                ExcelDocument.reader()
                        .sheet("TestSheet")
                        .headers(Arrays.asList())
                        .create()
        )
                .isInstanceOf(InvalidDocumentStructureException.class)
                .hasMessageContaining("must have at least one header");
    }

    @Test
    @DisplayName("ExcelDocument 빌더에서 행/열 불일치 시 InvalidDocumentStructureException 발생")
    void rowColumnMismatch() {
        assertThatThrownBy(() ->
                ExcelDocument.reader()
                        .sheet("TestSheet")
                        .headers(Arrays.asList("Header1", "Header2"))
                        .rows(Arrays.asList(
                                Arrays.asList("Value1")
                        ))
                        .create()
        )
                .isInstanceOf(InvalidDocumentStructureException.class)
                .hasMessageContaining("has 1 columns but expected 2");
    }

    @Test
    @DisplayName("중복 헤더명으로 문서 생성 시 HeaderMismatchException 발생")
    void duplicateHeadersInDocument() {
        assertThatThrownBy(() ->
                ExcelDocument.reader()
                        .sheet("TestSheet")
                        .headers(Arrays.asList("Header1", "Header1"))
                        .create()
        )
                .isInstanceOf(HeaderMismatchException.class)
                .hasMessageContaining("duplicate");
    }

    public static class ValidReadDto {
        @ExcelReadColumn(headerName = "Name")
        private String name;

        public ValidReadDto() {}
    }

    public static class ValidWriteDto {
        @ExcelWriteColumn(headerName = "Name", order = 1)
        private String name;

        public ValidWriteDto(String name) {
            this.name = name;
        }
    }

    public static class NoAnnotationDto {
        private String name;

        public NoAnnotationDto() {}
    }

    public static class DuplicateHeaderDto {
        @ExcelWriteColumn(headerName = "Name", order = 1)
        private String name1;

        @ExcelWriteColumn(headerName = "Name", order = 2)
        private String name2;

        public DuplicateHeaderDto(
                String name1,
                String name2
        ) {
            this.name1 = name1;
            this.name2 = name2;
        }
    }

    public static class EmptyHeaderDto {
        @ExcelWriteColumn(headerName = "", order = 1)
        private String name;

        public EmptyHeaderDto(String name) {
            this.name = name;
        }
    }

    private File createEmptyExcelFile() throws IOException {
        Path filePath = tempDir.resolve("empty.xlsx");
        Files.write(filePath, new byte[0]);
        return filePath.toFile();
    }

    private File createValidExcelFile() throws IOException {
        Path filePath = tempDir.resolve("valid.xlsx");

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("TestSheet");

            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Age");

            Row dataRow = sheet.createRow(1);
            dataRow.createCell(0).setCellValue("Hyunsoo");
            dataRow.createCell(1).setCellValue(30);

            try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
                workbook.write(fos);
            }
        }

        return filePath.toFile();
    }
}