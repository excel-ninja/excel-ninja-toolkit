package com.excelninja.infrastructure.persistence;

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
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

@DisplayName("NinjaExcelWriter Tests")
class NinjaExcelTest {

    static class UserWriteDto {
        @ExcelWriteColumn(headerName = "ID", order = 1)
        private final Long id;
        @ExcelWriteColumn(headerName = "Name", order = 2)
        private final String name;
        @ExcelWriteColumn(headerName = "Age", order = 3)
        private final Integer age;
        private final String ignore;

        public UserWriteDto(
                Long id,
                String name,
                Integer age
        ) {
            this.id = id;
            this.name = name;
            this.age = age;
            this.ignore = "This field should be ignored";
        }

    }


    public static class UserReadDto {
        @ExcelReadColumn(headerName = "ID")
        private Long id;
        @ExcelReadColumn(headerName = "Name")
        private  String name;
        @ExcelReadColumn(headerName = "Age")
        private  Integer age;

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
                Integer age
        ) {
            this.id = id;
            this.name = name;
            this.age = age;
        }

    }


    @Test
    @DisplayName("Write Excel via NinjaExcel")
    void write_and_read_excel_via_NinjaExcel() throws Exception {
        var usersToWrite = List.of(
                new UserWriteDto(1L, "Alice123#!@#!@3", 30),
                new UserWriteDto(2L, "Bob", 25)
        );

        var usersToRead = List.of(
                new UserReadDto(1L, "Alice123#!@#!@3", 30),
                new UserReadDto(2L, "Bob", 25)
        );



        var byteArrayOutputStream = new ByteArrayOutputStream();
        var document = ExcelDocument.createWriter(usersToWrite, "mySheetName");
        NinjaExcel.write(document, byteArrayOutputStream);
        var bytes = byteArrayOutputStream.toByteArray();
        assertTrue(bytes.length > 0);

        try (var wb = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            var sheet = wb.getSheet("mySheetName");
            assertNotNull(sheet);

            var header = sheet.getRow(0);
            assertEquals("ID", header.getCell(0).getStringCellValue());
            assertEquals("Name", header.getCell(1).getStringCellValue());
            assertEquals("Age", header.getCell(2).getStringCellValue());

            var row1 = sheet.getRow(1);
            assertEquals(1.0, row1.getCell(0).getNumericCellValue());
            assertEquals("Alice123#!@#!@3", row1.getCell(1).getStringCellValue());
            assertEquals(30.0, row1.getCell(2).getNumericCellValue());

            var row2 = sheet.getRow(2);
            assertEquals(2.0, row2.getCell(0).getNumericCellValue());
            assertEquals("Bob", row2.getCell(1).getStringCellValue());
            assertEquals(25.0, row2.getCell(2).getNumericCellValue());
        }

        var fileName = "/Users/hyunsoojo/users_test.xlsx";
        NinjaExcel.write(ExcelDocument.createWriter( usersToWrite), fileName);
        var path = Path.of(fileName);
        assertTrue(Files.exists(path));
        assertTrue(Files.size(path) > 0);

        Files.delete(path);
    }

    @Test
    @DisplayName("Read Excel via NinjaExcel")
    void read_excel_via_NinjaExcel() throws Exception {
        var file = new File(getClass().getClassLoader().getResource("users_test.xlsx").toURI());
        var users = NinjaExcel.read(file, UserReadDto.class);

        assertNotNull(users);
        assertEquals(2, users.size());

        var user1 = users.getFirst();
        assertEquals(1L, user1.id);
        assertEquals("Alice123#!@#!@3", user1.name);
        assertEquals(30, user1.age);

        var user2 = users.get(1);
        assertEquals(2L, user2.id);
        assertEquals("Bob", user2.name);
        assertEquals(25, user2.age);
    }

}