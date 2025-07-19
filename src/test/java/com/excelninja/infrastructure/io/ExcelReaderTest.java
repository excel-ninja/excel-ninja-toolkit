package com.excelninja.infrastructure.io;

import com.excelninja.application.facade.NinjaExcel;
import com.excelninja.domain.annotation.ExcelReadColumn;
import com.excelninja.domain.annotation.ExcelWriteColumn;
import com.excelninja.domain.model.ExcelDocument;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
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
        ExcelDocument document = ExcelDocument.writer().objects(testData).create();
        NinjaExcel.write(document, testFile.toString());

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
        ExcelDocument document = ExcelDocument.writer().objects(testData).create();
        NinjaExcel.write(document, testFile.toString());

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
}