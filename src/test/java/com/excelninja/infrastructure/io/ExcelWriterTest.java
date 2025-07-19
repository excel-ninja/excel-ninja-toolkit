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
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

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
        @ExcelWriteColumn(headerName = "Salary", order = 4)
        private final BigDecimal salary;
        @ExcelWriteColumn(headerName = "Hire Date", order = 5)
        private final LocalDate hireDate;
        @ExcelWriteColumn(headerName = "Last Login", order = 6)
        private final LocalDateTime lastLogin;
        @ExcelWriteColumn(headerName = "Is Active", order = 7)
        private final Boolean isActive;

        private final String ignore;

        public UserWriteDto(
                Long id,
                String name,
                Integer age,
                BigDecimal salary,
                LocalDate hireDate,
                LocalDateTime lastLogin,
                Boolean isActive
        ) {
            this.id = id;
            this.name = name;
            this.age = age;
            this.salary = salary;
            this.hireDate = hireDate;
            this.lastLogin = lastLogin;
            this.isActive = isActive;
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
        @ExcelReadColumn(headerName = "Salary")
        private BigDecimal salary;
        @ExcelReadColumn(headerName = "Hire Date")
        private LocalDate hireDate;
        @ExcelReadColumn(headerName = "Last Login")
        private LocalDateTime lastLogin;
        @ExcelReadColumn(headerName = "Is Active")
        private Boolean isActive;

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

        public void setSalary(BigDecimal salary) {
            this.salary = salary;
        }

        public void setHireDate(LocalDate hireDate) {
            this.hireDate = hireDate;
        }

        public void setLastLogin(LocalDateTime lastLogin) {
            this.lastLogin = lastLogin;
        }

        public void setIsActive(Boolean isActive) {
            this.isActive = isActive;
        }

        public UserReadDto(
                Long id,
                String name,
                Integer age,
                BigDecimal salary,
                LocalDate hireDate,
                LocalDateTime lastLogin,
                Boolean isActive
        ) {
            this.id = id;
            this.name = name;
            this.age = age;
            this.salary = salary;
            this.hireDate = hireDate;
            this.lastLogin = lastLogin;
            this.isActive = isActive;
        }
    }

    @Test
    @DisplayName("BigDecimal과 LocalDate/LocalDateTime을 포함한 엑셀 작성 및 읽기")
    void write_and_read_excel_with_enhanced_types() throws Exception {
        List<UserWriteDto> usersToWrite = Arrays.asList(
                new UserWriteDto(
                        1L,
                        "Alice Johnson",
                        30,
                        new BigDecimal("85000.50"),
                        LocalDate.of(2020, 3, 15),
                        LocalDateTime.of(2024, 1, 15, 9, 30, 0),
                        true
                ),
                new UserWriteDto(
                        2L,
                        "Bob Smith",
                        25,
                        new BigDecimal("72000.75"),
                        LocalDate.of(2021, 7, 20),
                        LocalDateTime.of(2024, 1, 14, 17, 45, 30),
                        false
                )
        );

        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        ExcelDocument document = ExcelDocument.writer().objects(usersToWrite).sheetName("EmployeeData").create();
        NinjaExcel.write(document, byteArrayOutputStream);
        byte[] bytes = byteArrayOutputStream.toByteArray();
        assertTrue(bytes.length > 0);

        try (org.apache.poi.ss.usermodel.Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("EmployeeData");
            assertNotNull(sheet);

            org.apache.poi.ss.usermodel.Row header = sheet.getRow(0);
            assertEquals("ID", header.getCell(0).getStringCellValue());
            assertEquals("Name", header.getCell(1).getStringCellValue());
            assertEquals("Age", header.getCell(2).getStringCellValue());
            assertEquals("Salary", header.getCell(3).getStringCellValue());
            assertEquals("Hire Date", header.getCell(4).getStringCellValue());
            assertEquals("Last Login", header.getCell(5).getStringCellValue());
            assertEquals("Is Active", header.getCell(6).getStringCellValue());

            org.apache.poi.ss.usermodel.Row row1 = sheet.getRow(1);
            assertEquals(1.0, row1.getCell(0).getNumericCellValue());
            assertEquals("Alice Johnson", row1.getCell(1).getStringCellValue());
            assertEquals(30.0, row1.getCell(2).getNumericCellValue());
            assertEquals(85000.50, row1.getCell(3).getNumericCellValue());
            assertNotNull(row1.getCell(4).getDateCellValue());
            assertNotNull(row1.getCell(5).getDateCellValue());
            assertTrue(row1.getCell(6).getBooleanCellValue());

            org.apache.poi.ss.usermodel.Row row2 = sheet.getRow(2);
            assertEquals(2.0, row2.getCell(0).getNumericCellValue());
            assertEquals("Bob Smith", row2.getCell(1).getStringCellValue());
            assertEquals(25.0, row2.getCell(2).getNumericCellValue());
            assertEquals(72000.75, row2.getCell(3).getNumericCellValue());
            assertNotNull(row2.getCell(4).getDateCellValue());
            assertNotNull(row2.getCell(5).getDateCellValue());
            assertFalse(row2.getCell(6).getBooleanCellValue());
        }

        Path tempFile = Files.createTempFile("enhanced_users_test", ".xlsx");
        try {
            ExcelDocument excelDocument = ExcelDocument.writer().objects(usersToWrite).create();
            NinjaExcel.write(excelDocument, tempFile.toString());
            assertTrue(Files.exists(tempFile));
            assertTrue(Files.size(tempFile) > 0);

            List<UserReadDto> readUsers = NinjaExcel.read(tempFile.toFile(), UserReadDto.class);
            assertNotNull(readUsers);
            assertEquals(2, readUsers.size());

            UserReadDto user1 = readUsers.get(0);
            assertEquals(1L, user1.id);
            assertEquals("Alice Johnson", user1.name);
            assertEquals(30, user1.age);
            assertEquals(new BigDecimal("85000.5"), user1.salary);
            assertEquals(LocalDate.of(2020, 3, 15), user1.hireDate);
            assertEquals(LocalDateTime.of(2024, 1, 15, 9, 30, 0), user1.lastLogin);
            assertTrue(user1.isActive);

            UserReadDto user2 = readUsers.get(1);
            assertEquals(2L, user2.id);
            assertEquals("Bob Smith", user2.name);
            assertEquals(25, user2.age);
            assertEquals(new BigDecimal("72000.75"), user2.salary);
            assertEquals(LocalDate.of(2021, 7, 20), user2.hireDate);
            assertEquals(LocalDateTime.of(2024, 1, 14, 17, 45, 30), user2.lastLogin);
            assertFalse(user2.isActive);
        } finally {
            Files.deleteIfExists(tempFile);
        }
    }

    @Test
    @DisplayName("BigDecimal 정밀도 테스트")
    void bigDecimal_precision_test() throws Exception {
        List<UserWriteDto> usersWithPreciseDecimals = Collections.singletonList(
                new UserWriteDto(
                        1L,
                        "Test User",
                        30,
                        new BigDecimal("12345.6789"),
                        LocalDate.of(2024, 1, 1),
                        LocalDateTime.of(2024, 1, 1, 12, 0, 0),
                        true
                )
        );

        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        ExcelDocument document = ExcelDocument.writer().objects(usersWithPreciseDecimals).sheetName("PrecisionTest").create();
        NinjaExcel.write(document, byteArrayOutputStream);
        byte[] bytes = byteArrayOutputStream.toByteArray();

        try (org.apache.poi.ss.usermodel.Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("PrecisionTest");
            org.apache.poi.ss.usermodel.Row row1 = sheet.getRow(1);
            assertEquals(12345.6789, row1.getCell(3).getNumericCellValue(), 0.0001);
        }
    }

    @Test
    @DisplayName("날짜/시간 포맷 확인 테스트")
    void date_time_formatting_test() throws Exception {
        List<UserWriteDto> usersWithDates = Collections.singletonList(
                new UserWriteDto(
                        1L,
                        "Date Test User",
                        25,
                        new BigDecimal("50000.00"),
                        LocalDate.of(2023, 12, 25),
                        LocalDateTime.of(2024, 1, 1, 0, 0, 0),
                        true
                )
        );

        Path tempFile = Files.createTempFile("date_test", ".xlsx");
        try {
            ExcelDocument document = ExcelDocument.writer().objects(usersWithDates).sheetName("DateTest").create();
            NinjaExcel.write(document, tempFile.toString());

            List<UserReadDto> readUsers = NinjaExcel.read(tempFile.toFile(), UserReadDto.class);
            UserReadDto user = readUsers.get(0);

            assertEquals(LocalDate.of(2023, 12, 25), user.hireDate);
            assertEquals(LocalDateTime.of(2024, 1, 1, 0, 0, 0), user.lastLogin);

        } finally {
            Files.deleteIfExists(tempFile);
        }
    }
}