package com.excelninja.infrastructure.io;

import com.excelninja.application.facade.NinjaExcel;
import com.excelninja.domain.annotation.ExcelReadColumn;
import com.excelninja.domain.annotation.ExcelWriteColumn;
import com.excelninja.domain.model.ExcelDocument;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

import static org.junit.jupiter.api.Assertions.*;

@DisplayName("엑셀 작성기 테스트")
class ExcelWriterTest {

    static class UserWriteDto {
        @ExcelWriteColumn(headerName = "ID", order = 1)
        private final Long id;
        @ExcelWriteColumn(headerName = "Name", order = 2)
        private final String name;
        @ExcelWriteColumn(headerName = "Age", order = 3)
        private final Integer age;
        @ExcelWriteColumn(headerName = "Score", order = 4)
        private final Double score;

        private final String ignore;

        public UserWriteDto(
                Long id,
                String name,
                Integer age,
                Double score
        ) {
            this.id = id;
            this.name = name;
            this.age = age;
            this.score = score;
            this.ignore = "This field should be ignored";
        }
    }

    public static class UserReadDto {
        @ExcelReadColumn(headerName = "ID")
        private Long id;
        @ExcelReadColumn(headerName = "Name")
        private String name;
        @ExcelReadColumn(headerName = "Age")
        private Integer age;
        @ExcelWriteColumn(headerName = "Score")
        private Double score;

        public UserReadDto() {}

        public void setId(Long id) {
            this.id = id;
        }

        public void setName(String name) {
            this.name = name;
        }

        public void setAge(Integer age) {
            this.age = age;
        }

        public UserReadDto(
                Long id,
                String name,
                Integer age,
                Double score
        ) {
            this.id = id;
            this.name = name;
            this.age = age;
            this.score = score;
        }
    }

    @Test
    @DisplayName("임시파일로 엑셀 작성 및 읽기")
    void write_and_read_excel_via_NinjaExcel() throws Exception {
        List<UserWriteDto> usersToWrite = Arrays.asList(
                new UserWriteDto(1L, "Alice123#!@#!@3", 30, 95.5),
                new UserWriteDto(2L, "Bob", 25, 88.0)
        );

        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        ExcelDocument document = ExcelDocument.createFromEntities(usersToWrite, "mySheetName");
        NinjaExcel.write(document, byteArrayOutputStream);
        byte[] bytes = byteArrayOutputStream.toByteArray();
        assertTrue(bytes.length > 0);

        try (org.apache.poi.ss.usermodel.Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("mySheetName");
            assertNotNull(sheet);

            org.apache.poi.ss.usermodel.Row header = sheet.getRow(0);
            assertEquals("ID", header.getCell(0).getStringCellValue());
            assertEquals("Name", header.getCell(1).getStringCellValue());
            assertEquals("Age", header.getCell(2).getStringCellValue());

            org.apache.poi.ss.usermodel.Row row1 = sheet.getRow(1);
            assertEquals(1.0, row1.getCell(0).getNumericCellValue());
            assertEquals("Alice123#!@#!@3", row1.getCell(1).getStringCellValue());
            assertEquals(30.0, row1.getCell(2).getNumericCellValue());

            org.apache.poi.ss.usermodel.Row row2 = sheet.getRow(2);
            assertEquals(2.0, row2.getCell(0).getNumericCellValue());
            assertEquals("Bob", row2.getCell(1).getStringCellValue());
            assertEquals(25.0, row2.getCell(2).getNumericCellValue());
        }

        Path tempFile = Files.createTempFile("users_test", ".xlsx");
        try {
            NinjaExcel.write(ExcelDocument.createFromEntities(usersToWrite), tempFile.toString());
            assertTrue(Files.exists(tempFile));
            assertTrue(Files.size(tempFile) > 0);
        } finally {
            Files.deleteIfExists(tempFile);
        }
    }

    @Test
    @DisplayName("엑셀 파일 읽기")
    void read_excel_via_NinjaExcel() throws Exception {
        File file = new File(Objects.requireNonNull(getClass().getClassLoader().getResource("users_test.xlsx")).toURI());
        List<UserReadDto> users = NinjaExcel.read(file, UserReadDto.class);

        assertNotNull(users);
        assertEquals(2, users.size());

        UserReadDto user1 = users.get(0);
        assertEquals(1L, user1.id);
        assertEquals("Alice123#!@#!@3", user1.name);
        assertEquals(30, user1.age);

        UserReadDto user2 = users.get(1);
        assertEquals(2L, user2.id);
        assertEquals("Bob", user2.name);
        assertEquals(25, user2.age);
    }

}