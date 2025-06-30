package com.excelninja.application;

import com.excelninja.application.facade.NinjaExcel;
import com.excelninja.domain.annotation.ExcelColumn;
import com.excelninja.domain.model.ExcelDocument;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

@DisplayName("NinjaExcelWriter Tests")
class NinjaExcelWriterTest {

    static class UserDto {
        @ExcelColumn(headerName = "ID", order = 1)
        private final Long id;
        @ExcelColumn(headerName = "Name", order = 2)
        private final String name;
        @ExcelColumn(headerName = "Age", order = 3)
        private final Integer age;
        private final String ignore;

        public UserDto(
                Long id,
                String name,
                Integer age
        ) {
            this.id = id;
            this.name = name;
            this.age = age;
            this.ignore = "This field should be ignored";
        }

        public Long getId() {return id;}

        public String getName() {return name;}

        public Integer getAge() {return age;}

        public String getIgnore() {return ignore;}
    }

    @Test
    @DisplayName("Write Excel via NinjaExcel")
    void write_and_read_excel_via_NinjaExcel() throws Exception {
        var users = List.of(
                new UserDto(1L, "Alice123#!@#!@3", 30),
                new UserDto(2L, "Bob", 25)
        );

        var byteArrayOutputStream = new ByteArrayOutputStream();
        NinjaExcel.write(ExcelDocument.fromDtoList(users,"mySheetName"), byteArrayOutputStream);
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
        NinjaExcel.write(ExcelDocument.fromDtoList( users), fileName);
        var path = Path.of(fileName);
        assertTrue(Files.exists(path));
        assertTrue(Files.size(path) > 0);

//        Files.delete(path);
    }
}