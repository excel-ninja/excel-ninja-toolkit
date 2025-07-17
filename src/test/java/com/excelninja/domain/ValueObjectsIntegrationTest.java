package com.excelninja.domain;

import com.excelninja.domain.annotation.ExcelReadColumn;
import com.excelninja.domain.annotation.ExcelWriteColumn;
import com.excelninja.domain.exception.HeaderMismatchException;
import com.excelninja.domain.exception.InvalidDocumentStructureException;
import com.excelninja.domain.model.*;
import com.excelninja.infrastructure.converter.DefaultConverter;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

@DisplayName("값 객체 통합 테스트")
class ValueObjectsIntegrationTest {

    @Test
    @DisplayName("값 객체를 사용한 완전한 Excel 처리 플로우")
    void completeExcelProcessingWithValueObjects() {
        SheetName sheetName = new SheetName("UserData");
        Headers headers = Headers.of("Name", "Age", "Email", "Active");
        DocumentRows rows = DocumentRows.of(Arrays.asList(
                Arrays.asList("Hyunsoo", 30, "hyunsoo@example.com", true),
                Arrays.asList("Eunmi", 25, "eunmi@example.com", false),
                Arrays.asList("naruto", 35, "naruto@example.com", true)
        ), 4);

        ExcelDocument document = ExcelDocument.readBuilder()
                .sheet(sheetName)
                .headers(headers)
                .rows(rows)
                .build();

        List<UserDto> users = document.convertToEntities(UserDto.class, new DefaultConverter());

        assertThat(users).hasSize(3);

        UserDto firstUser = users.get(0);
        assertThat(firstUser.getName()).isEqualTo("Hyunsoo");
        assertThat(firstUser.getAge()).isEqualTo(30);
        assertThat(firstUser.getEmail()).isEqualTo("hyunsoo@example.com");
        assertThat(firstUser.isActive()).isTrue();

        assertThat(document.getSheetName()).isInstanceOf(SheetName.class);
        assertThat(document.getHeaders()).isInstanceOf(Headers.class);
        assertThat(document.getRows()).isInstanceOf(DocumentRows.class);

        assertThat(document.getCellValue(0, "Name")).isEqualTo("Hyunsoo");
        assertThat(document.getCellValue(1, "Age", Integer.class)).isEqualTo(25);
        assertThat(document.getColumn("Email")).containsExactly("hyunsoo@example.com", "eunmi@example.com", "naruto@example.com");
    }

    @Test
    @DisplayName("값 객체를 사용한 엔티티에서 Excel 문서 생성")
    void createExcelDocumentFromEntitiesWithValueObjects() {
        List<UserDto> users = Arrays.asList(
                new UserDto("Alice", 28, "alice@example.com", true),
                new UserDto("Charlie", 32, "charlie@example.com", false)
        );

        ExcelDocument document = ExcelDocument.writeBuilder().objects(users).sheetName("UserReport").build();

        assertThat(document.getSheetName().getValue()).isEqualTo("UserReport");
        assertThat(document.getSheetName().isDefault()).isFalse();

        Headers headers = document.getHeaders();
        assertThat(headers.size()).isEqualTo(4);
        assertThat(headers.containsHeader("Name")).isTrue();
        assertThat(headers.containsHeader("Age")).isTrue();
        assertThat(headers.containsHeader("Email")).isTrue();
        assertThat(headers.containsHeader("Active")).isTrue();

        DocumentRows rows = document.getRows();
        assertThat(rows.size()).isEqualTo(2);
        assertThat(rows.getExpectedColumnCount()).isEqualTo(4);

        DocumentRow firstRow = rows.getRow(0);
        assertThat(firstRow.getValueByHeader(headers, "Name")).isEqualTo("Alice");
        assertThat(firstRow.getValueByHeader(headers, "Age", Integer.class)).isEqualTo(28);
        assertThat(firstRow.getValueByHeader(headers, "Email")).isEqualTo("alice@example.com");
        assertThat(firstRow.getValueByHeader(headers, "Active", Boolean.class)).isTrue();
    }

    @Test
    @DisplayName("값 객체 도메인 규칙 검증")
    void valueObjectDomainRulesValidation() {
        assertThatThrownBy(() -> new SheetName(""))
                .isInstanceOf(InvalidDocumentStructureException.class)
                .hasMessageContaining("Sheet name cannot be null or empty");

        assertThatThrownBy(() -> new SheetName("Sheet/With/Slash"))
                .isInstanceOf(InvalidDocumentStructureException.class)
                .hasMessageContaining("contains invalid character");

        assertThatThrownBy(() -> new SheetName("History"))
                .isInstanceOf(InvalidDocumentStructureException.class)
                .hasMessageContaining("reserved and cannot be used");

        assertThatThrownBy(() -> Headers.of("Name", "Name"))
                .isInstanceOf(HeaderMismatchException.class)
                .hasMessageContaining("duplicate");

        assertThatThrownBy(() -> Headers.of())
                .isInstanceOf(InvalidDocumentStructureException.class)
                .hasMessageContaining("must have at least one header");

        assertThatThrownBy(() -> new DocumentRow(Arrays.asList("value"), -1))
                .isInstanceOf(InvalidDocumentStructureException.class)
                .hasMessageContaining("Row number cannot be negative");
    }

    @Test
    @DisplayName("값 객체를 사용한 Excel 파일 읽기/쓰기 (NinjaExcel 통합)")
    void excelFileReadWriteWithValueObjects() {
        ExcelDocument originalDocument = ExcelDocument.readBuilder()
                .sheet("Integration Test")
                .headers("ID", "Name", "Score")
                .rows(Arrays.asList(
                        Arrays.asList(1L, "Test User 1", 95.5),
                        Arrays.asList(2L, "Test User 2", 87.0)
                ))
                .build();

        List<TestScoreDto> scores = originalDocument.convertToEntities(TestScoreDto.class, new DefaultConverter());

        assertThat(scores).hasSize(2);

        TestScoreDto firstScore = scores.get(0);
        assertThat(firstScore.getId()).isEqualTo(1L);
        assertThat(firstScore.getName()).isEqualTo("Test User 1");
        assertThat(firstScore.getScore()).isEqualTo(95.5);

        ExcelDocument recreatedDocument = ExcelDocument.writeBuilder().objects(scores).sheetName("Recreated").build();

        assertThat(recreatedDocument.getSheetName().getValue()).isEqualTo("Recreated");
        assertThat(recreatedDocument.getHeaders().size()).isEqualTo(3);
        assertThat(recreatedDocument.getRows().size()).isEqualTo(2);

        assertThat(recreatedDocument.getCellValue(0, "ID")).isEqualTo(1L);
        assertThat(recreatedDocument.getCellValue(1, "Name")).isEqualTo("Test User 2");
        assertThat(recreatedDocument.getCellValue(1, "Score")).isEqualTo(87.0);
    }

    @Test
    @DisplayName("값 객체의 불변성 보장")
    void valueObjectImmutability() {
        SheetName sheetName = new SheetName("ImmutableTest");
        Headers headers = Headers.of("Col1", "Col2", "Col3");
        DocumentRows rows = DocumentRows.of(Arrays.asList(
                Arrays.asList("A", "B", "C"),
                Arrays.asList("X", "Y", "Z")
        ), 3);

        ExcelDocument document = ExcelDocument.readBuilder()
                .sheet(sheetName)
                .headers(headers)
                .rows(rows)
                .build();

        Headers retrievedHeaders = document.getHeaders();
        DocumentRows retrievedRows = document.getRows();

        assertThatThrownBy(() -> retrievedHeaders.getHeaders().add(new Header("NewHeader", 3)))
                .isInstanceOf(UnsupportedOperationException.class);

        assertThatThrownBy(() -> retrievedRows.getRows().add(DocumentRow.of(Arrays.asList("New", "Row", "Data"), 3)))
                .isInstanceOf(UnsupportedOperationException.class);

        SheetName originalSheetName = document.getSheetName();
        assertThat(originalSheetName.getValue()).isEqualTo("ImmutableTest");

        SheetName newSheetName = new SheetName("NewSheet");
        assertThat(originalSheetName.getValue()).isEqualTo("ImmutableTest"); // 변경되지 않음
        assertThat(newSheetName.getValue()).isEqualTo("NewSheet");
    }

    @Test
    @DisplayName("값 객체를 사용한 도메인 질의 메서드")
    void valueObjectDomainQueries() {
        ExcelDocument document = ExcelDocument.readBuilder()
                .sheet("QueryTest")
                .headers("Product", "Category", "Price", "InStock")
                .rows(Arrays.asList(
                        Arrays.asList("Laptop", "Electronics", 999.99, true),
                        Arrays.asList("Desk", "Furniture", 299.99, false),
                        Arrays.asList("Phone", "Electronics", 599.99, true),
                        Arrays.asList("Chair", "Furniture", 149.99, true)
                ))
                .build();

        assertThat(document.hasData()).isTrue();
        assertThat(document.getRowCount()).isEqualTo(4);
        assertThat(document.getColumnCount()).isEqualTo(4);

        List<Object> categories = document.getColumn("Category");
        assertThat(categories).containsExactly("Electronics", "Furniture", "Electronics", "Furniture");

        assertThat(document.getCellValue(0, "Product")).isEqualTo("Laptop");
        assertThat(document.getCellValue(2, "Price", Double.class)).isEqualTo(599.99);
        assertThat(document.getCellValue(3, "InStock", Boolean.class)).isTrue();

        Headers headers = document.getHeaders();
        assertThat(headers.containsHeader("Price")).isTrue();
        assertThat(headers.containsHeader("Weight")).isFalse();
        assertThat(headers.getPositionOf("InStock")).isEqualTo(3);

        DocumentRows rows = document.getRows();
        DocumentRow firstRow = rows.getRow(0);
        assertThat(firstRow.getRowNumber()).isEqualTo(1);
        assertThat(firstRow.isEmpty()).isFalse();
        assertThat(firstRow.hasValue(0)).isTrue();
        assertThat(firstRow.hasValue(10)).isFalse();
    }

    public static class UserDto {
        @ExcelReadColumn(headerName = "Name")
        @ExcelWriteColumn(headerName = "Name", order = 1)
        private String name;

        @ExcelReadColumn(headerName = "Age")
        @ExcelWriteColumn(headerName = "Age", order = 2)
        private Integer age;

        @ExcelReadColumn(headerName = "Email")
        @ExcelWriteColumn(headerName = "Email", order = 3)
        private String email;

        @ExcelReadColumn(headerName = "Active")
        @ExcelWriteColumn(headerName = "Active", order = 4)
        private Boolean active;

        public UserDto() {}

        public UserDto(
                String name,
                Integer age,
                String email,
                Boolean active
        ) {
            this.name = name;
            this.age = age;
            this.email = email;
            this.active = active;
        }

        public String getName() {return name;}

        public void setName(String name) {this.name = name;}

        public Integer getAge() {return age;}

        public void setAge(Integer age) {this.age = age;}

        public String getEmail() {return email;}

        public void setEmail(String email) {this.email = email;}

        public Boolean isActive() {return active;}

        public void setActive(Boolean active) {this.active = active;}
    }

    public static class TestScoreDto {
        @ExcelReadColumn(headerName = "ID")
        @ExcelWriteColumn(headerName = "ID", order = 1)
        private Long id;

        @ExcelReadColumn(headerName = "Name")
        @ExcelWriteColumn(headerName = "Name", order = 2)
        private String name;

        @ExcelReadColumn(headerName = "Score")
        @ExcelWriteColumn(headerName = "Score", order = 3)
        private Double score;

        public TestScoreDto() {}

        public TestScoreDto(
                Long id,
                String name,
                Double score
        ) {
            this.id = id;
            this.name = name;
            this.score = score;
        }

        public Long getId() {return id;}

        public void setId(Long id) {this.id = id;}

        public String getName() {return name;}

        public void setName(String name) {this.name = name;}

        public Double getScore() {return score;}

        public void setScore(Double score) {this.score = score;}
    }
}
