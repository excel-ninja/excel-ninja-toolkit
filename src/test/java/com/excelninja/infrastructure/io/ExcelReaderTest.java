package com.excelninja.infrastructure.io;

import com.excelninja.application.facade.NinjaExcel;
import com.excelninja.domain.annotation.ExcelReadColumn;
import com.excelninja.domain.annotation.ExcelWriteColumn;
import com.excelninja.domain.model.ExcelWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.OutputStream;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import static org.assertj.core.api.AssertionsForInterfaceTypes.assertThat;

class ExcelReaderTest {

    @TempDir
    Path tempDir;

    public static class UserTestDto {
        @ExcelReadColumn(headerName = "ID")
        @ExcelWriteColumn(headerName = "ID", order = 1)
        private Long id;

        @ExcelReadColumn(headerName = "Name")
        @ExcelWriteColumn(headerName = "Name", order = 2)
        private String name;

        @ExcelReadColumn(headerName = "Age")
        @ExcelWriteColumn(headerName = "Age", order = 3)
        private Integer age;

        @ExcelReadColumn(headerName = "Salary")
        @ExcelWriteColumn(headerName = "Salary", order = 4)
        private BigDecimal salary;

        @ExcelReadColumn(headerName = "birthday")
        @ExcelWriteColumn(headerName = "birthday", order = 5)
        private LocalDate birthday;

        @ExcelReadColumn(headerName = "lastUpdated")
        @ExcelWriteColumn(headerName = "lastUpdated", order = 6)
        private LocalDateTime lastUpdated;

        public UserTestDto() {}

        public UserTestDto(
                Long id,
                String name,
                Integer age,
                BigDecimal salary,
                LocalDate birthday,
                LocalDateTime lastUpdated
        ) {
            this.id = id;
            this.name = name;
            this.age = age;
            this.salary = salary;
            this.birthday = birthday;
            this.lastUpdated = lastUpdated;
        }
    }

    public static class DateTargetWriteDto {
        @ExcelWriteColumn(headerName = "DateAsDate", order = 1)
        private final LocalDate dateAsDate;

        @ExcelWriteColumn(headerName = "DateAsString", order = 2)
        private final LocalDate dateAsString;

        @ExcelWriteColumn(headerName = "DateAsLocalDate", order = 3)
        private final LocalDate dateAsLocalDate;

        @ExcelWriteColumn(headerName = "DateTimeAsString", order = 4)
        private final LocalDateTime dateTimeAsString;

        @ExcelWriteColumn(headerName = "DateTimeAsDate", order = 5)
        private final LocalDateTime dateTimeAsDate;

        @ExcelWriteColumn(headerName = "DateTimeAsLocalDateTime", order = 6)
        private final LocalDateTime dateTimeAsLocalDateTime;

        public DateTargetWriteDto(
                LocalDate dateAsDate,
                LocalDate dateAsString,
                LocalDate dateAsLocalDate,
                LocalDateTime dateTimeAsString,
                LocalDateTime dateTimeAsDate,
                LocalDateTime dateTimeAsLocalDateTime
        ) {
            this.dateAsDate = dateAsDate;
            this.dateAsString = dateAsString;
            this.dateAsLocalDate = dateAsLocalDate;
            this.dateTimeAsString = dateTimeAsString;
            this.dateTimeAsDate = dateTimeAsDate;
            this.dateTimeAsLocalDateTime = dateTimeAsLocalDateTime;
        }
    }

    public static class DateTargetReadDto {
        @ExcelReadColumn(headerName = "DateAsDate")
        private Date dateAsDate;

        @ExcelReadColumn(headerName = "DateAsString")
        private String dateAsString;

        @ExcelReadColumn(headerName = "DateAsLocalDate")
        private LocalDate dateAsLocalDate;

        @ExcelReadColumn(headerName = "DateTimeAsString")
        private String dateTimeAsString;

        @ExcelReadColumn(headerName = "DateTimeAsDate")
        private Date dateTimeAsDate;

        @ExcelReadColumn(headerName = "DateTimeAsLocalDateTime")
        private LocalDateTime dateTimeAsLocalDateTime;
    }

    public static class NullableScalarWriteDto {
        @ExcelWriteColumn(headerName = "ID", order = 1)
        private final Long id;

        @ExcelWriteColumn(headerName = "Age", order = 2)
        private final Integer age;

        @ExcelWriteColumn(headerName = "Salary", order = 3)
        private final BigDecimal salary;

        @ExcelWriteColumn(headerName = "Active", order = 4)
        private final Boolean active;

        @ExcelWriteColumn(headerName = "Birthday", order = 5)
        private final LocalDate birthday;

        @ExcelWriteColumn(headerName = "LastUpdated", order = 6)
        private final LocalDateTime lastUpdated;

        public NullableScalarWriteDto(
                Long id,
                Integer age,
                BigDecimal salary,
                Boolean active,
                LocalDate birthday,
                LocalDateTime lastUpdated
        ) {
            this.id = id;
            this.age = age;
            this.salary = salary;
            this.active = active;
            this.birthday = birthday;
            this.lastUpdated = lastUpdated;
        }
    }

    public static class NullableScalarReadDto {
        @ExcelReadColumn(headerName = "ID")
        private Long id;

        @ExcelReadColumn(headerName = "Age")
        private Integer age;

        @ExcelReadColumn(headerName = "Salary")
        private BigDecimal salary;

        @ExcelReadColumn(headerName = "Active")
        private Boolean active;

        @ExcelReadColumn(headerName = "Birthday")
        private LocalDate birthday;

        @ExcelReadColumn(headerName = "LastUpdated")
        private LocalDateTime lastUpdated;
    }

    public static class StringCellReadDto {
        @ExcelReadColumn(headerName = "Age")
        private Integer age;

        @ExcelReadColumn(headerName = "Salary", defaultValue = "0.5")
        private Double salary;

        @ExcelReadColumn(headerName = "Active")
        private Boolean active;
    }

    @Test
    @DisplayName("엑셀 파일에서 유효한 데이터 쓰고 읽기")
    void readValidExcelFile() {
        List<UserTestDto> testData = Arrays.asList(
                new UserTestDto(
                        1L,
                        "Alice123#!@#!@3",
                        30,
                        new BigDecimal("100000.0"),
                        LocalDate.of(1993, 1, 1),
                        LocalDateTime.of(2024, 1, 15, 10, 30, 0)
                )
        );

        Path testFile = tempDir.resolve("users_test.xlsx");
        ExcelWorkbook workbook = ExcelWorkbook.builder().sheet(testData).build();
        NinjaExcel.write(workbook, testFile.toString());

        List<UserTestDto> readDTO = NinjaExcel.read(testFile.toFile(), UserTestDto.class);

        assertThat(readDTO).hasSize(1);

        UserTestDto dto = readDTO.get(0);
        assertThat(dto).isNotNull();
        assertThat(dto.id).isEqualTo(1L);
        assertThat(dto.name).isEqualTo("Alice123#!@#!@3");
        assertThat(dto.age).isEqualTo(30);
        assertThat(dto.salary).isEqualTo(new BigDecimal("100000.0"));
        assertThat(dto.birthday).isEqualTo(LocalDate.of(1993, 1, 1));
        assertThat(dto.lastUpdated).isEqualTo(LocalDateTime.of(2024, 1, 15, 10, 30, 0));
    }

    @Test
    @DisplayName("엑셀 파일에서 여러 사용자 데이터 쓰고 읽기")
    void readExcelFileWithMultipleUsers() {
        List<UserTestDto> testData = Arrays.asList(
                new UserTestDto(
                        1L,
                        "Hyunsoo",
                        30,
                        new BigDecimal("100000.0"),
                        LocalDate.of(1993, 1, 1),
                        LocalDateTime.of(2024, 1, 15, 10, 30, 0)
                ),
                new UserTestDto(
                        2L,
                        "Eunmi",
                        25,
                        new BigDecimal("85000.5"),
                        LocalDate.of(1998, 5, 15),
                        LocalDateTime.of(2024, 1, 16, 14, 20, 0)
                ),
                new UserTestDto(
                        3L,
                        "Younkyung",
                        35,
                        new BigDecimal("120000.75"),
                        LocalDate.of(1988, 12, 10),
                        LocalDateTime.of(2024, 1, 17, 9, 15, 0)
                )
        );

        Path testFile = tempDir.resolve("multiple_users_test.xlsx");
        ExcelWorkbook workbook = ExcelWorkbook.builder().sheet(testData).build();
        NinjaExcel.write(workbook, testFile.toString());

        List<UserTestDto> readUsers = NinjaExcel.read(testFile.toFile(), UserTestDto.class);

        assertThat(readUsers).hasSize(3);

        UserTestDto alice = readUsers.get(0);
        assertThat(alice.id).isEqualTo(1L);
        assertThat(alice.name).isEqualTo("Hyunsoo");
        assertThat(alice.age).isEqualTo(30);
        assertThat(alice.salary).isEqualTo(new BigDecimal("100000.0"));
        assertThat(alice.birthday).isEqualTo(LocalDate.of(1993, 1, 1));
        assertThat(alice.lastUpdated).isEqualTo(LocalDateTime.of(2024, 1, 15, 10, 30, 0));

        UserTestDto bob = readUsers.get(1);
        assertThat(bob.id).isEqualTo(2L);
        assertThat(bob.name).isEqualTo("Eunmi");
        assertThat(bob.salary).isEqualTo(new BigDecimal("85000.5"));
        assertThat(bob.birthday).isEqualTo(LocalDate.of(1998, 5, 15));
        assertThat(bob.lastUpdated).isEqualTo(LocalDateTime.of(2024, 1, 16, 14, 20, 0));

        UserTestDto charlie = readUsers.get(2);
        assertThat(charlie.id).isEqualTo(3L);
        assertThat(charlie.name).isEqualTo("Younkyung");
        assertThat(charlie.age).isEqualTo(35);
        assertThat(charlie.salary).isEqualTo(new BigDecimal("120000.75"));
        assertThat(charlie.birthday).isEqualTo(LocalDate.of(1988, 12, 10));
        assertThat(charlie.lastUpdated).isEqualTo(LocalDateTime.of(2024, 1, 17, 9, 15, 0));
    }

    @Test
    @DisplayName("Date cells can be mapped to Date, String, LocalDate and LocalDateTime")
    void readDateCellsIntoMultipleTargetTypes() {
        LocalDate dateOnly = LocalDate.of(2024, 12, 25);
        LocalDateTime dateTime = LocalDateTime.of(2024, 12, 25, 14, 30, 5);
        List<DateTargetWriteDto> testData = Arrays.asList(
                new DateTargetWriteDto(dateOnly, dateOnly, dateOnly, dateTime, dateTime, dateTime)
        );

        Path testFile = tempDir.resolve("date_target_types.xlsx");
        ExcelWorkbook workbook = ExcelWorkbook.builder().sheet(testData).build();
        NinjaExcel.write(workbook, testFile.toString());

        List<DateTargetReadDto> rows = NinjaExcel.read(testFile.toFile(), DateTargetReadDto.class);

        assertThat(rows).hasSize(1);

        DateTargetReadDto row = rows.get(0);
        ZoneId zoneId = ZoneId.systemDefault();

        assertThat(row.dateAsDate.toInstant().atZone(zoneId).toLocalDate()).isEqualTo(dateOnly);
        assertThat(row.dateAsString).isEqualTo("2024-12-25");
        assertThat(row.dateAsLocalDate).isEqualTo(dateOnly);
        assertThat(row.dateTimeAsString).isEqualTo("2024-12-25T14:30:05");
        assertThat(row.dateTimeAsDate.toInstant().atZone(zoneId).toLocalDateTime()).isEqualTo(dateTime);
        assertThat(row.dateTimeAsLocalDateTime).isEqualTo(dateTime);
    }

    @Test
    @DisplayName("Null scalar values round-trip without becoming invalid empty strings")
    void readNullScalarCellsAsNull() {
        List<NullableScalarWriteDto> testData = Collections.singletonList(
                new NullableScalarWriteDto(1L, null, null, null, null, null)
        );

        Path testFile = tempDir.resolve("nullable_scalar_cells.xlsx");
        ExcelWorkbook workbook = ExcelWorkbook.builder().sheet(testData).build();
        NinjaExcel.write(workbook, testFile.toString());

        List<NullableScalarReadDto> rows = NinjaExcel.read(testFile.toFile(), NullableScalarReadDto.class);

        assertThat(rows).hasSize(1);
        NullableScalarReadDto row = rows.get(0);
        assertThat(row.id).isEqualTo(1L);
        assertThat(row.age).isNull();
        assertThat(row.salary).isNull();
        assertThat(row.active).isNull();
        assertThat(row.birthday).isNull();
        assertThat(row.lastUpdated).isNull();
    }

    @Test
    @DisplayName("POI reader converts string scalar cells and applies default values for empty strings")
    void readStringScalarCellsAndApplyDefaults() throws Exception {
        Path workbookPath = tempDir.resolve("string_scalar_cells.xlsx");

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Age");
            headerRow.createCell(1).setCellValue("Salary");
            headerRow.createCell(2).setCellValue("Active");

            Row dataRow = sheet.createRow(1);
            dataRow.createCell(0).setCellValue("42");
            dataRow.createCell(1).setCellValue("");
            dataRow.createCell(2).setCellValue("yes");

            try (OutputStream outputStream = java.nio.file.Files.newOutputStream(workbookPath)) {
                workbook.write(outputStream);
            }
        }

        List<StringCellReadDto> rows = NinjaExcel.read(workbookPath.toFile(), StringCellReadDto.class);

        assertThat(rows).hasSize(1);
        StringCellReadDto row = rows.get(0);
        assertThat(row.age).isEqualTo(42);
        assertThat(row.salary).isEqualTo(0.5d);
        assertThat(row.active).isEqualTo(true);
    }
}
