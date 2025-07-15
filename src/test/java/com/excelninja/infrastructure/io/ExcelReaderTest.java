package com.excelninja.infrastructure.io;

import com.excelninja.application.facade.NinjaExcel;
import com.excelninja.domain.annotation.ExcelReadColumn;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.net.URISyntaxException;
import java.util.List;

import static org.assertj.core.api.AssertionsForInterfaceTypes.assertThat;

class ExcelReaderTest {

    public static class UserReadDto {
        @ExcelReadColumn(headerName = "ID")
        private Long id;
        @ExcelReadColumn(headerName = "Name")
        private String name;
        @ExcelReadColumn(headerName = "Age")
        private Integer age;

        public UserReadDto() {}
    }

    @Test
    @DisplayName("Read valid Excel file and return correct document")
    void readValidExcelFile() throws URISyntaxException {

        File file = new File(getClass().getClassLoader().getResource("users_test.xlsx").toURI());
        List<UserReadDto> readDTO = NinjaExcel.read(file, UserReadDto.class);

        UserReadDto dto = readDTO.get(0);
        assertThat(dto).isNotNull();
        assertThat(dto.id).isEqualTo(1L);
        assertThat(dto.name).isEqualTo("Alice123#!@#!@3");
        assertThat(dto.age).isEqualTo(30);

        UserReadDto dto2 = readDTO.get(1);
        assertThat(dto2).isNotNull();
        assertThat(dto2.id).isEqualTo(2L);
        assertThat(dto2.name).isEqualTo("Bob");
        assertThat(dto2.age).isEqualTo(25);

        assertThat(readDTO).hasSize(2);
    }
}