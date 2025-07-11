package com.excelninja.infrastructure.persistence;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.util.List;

import static org.assertj.core.api.AssertionsForInterfaceTypes.assertThat;

class PoiExcelReaderTest {
    @Test
    @DisplayName("Read valid Excel file and return correct document")
    void readValidExcelFile() throws IOException, URISyntaxException {
        var file = new File(getClass().getClassLoader().getResource("users_test.xlsx").toURI());
        var reader = new PoiExcelReader();
        var document = reader.read(file);

        assertThat(document).isNotNull();
        assertThat(document.getSheetName()).isEqualTo("UserDto");
        assertThat(document.getHeaders()).containsExactly("ID", "Name", "Age");
        assertThat(document.getRows()).containsExactly(
                List.of(1.0, "Alice123#!@#!@3", 30.0),
                List.of(2.0, "Bob", 25.0)
        );
    }
}