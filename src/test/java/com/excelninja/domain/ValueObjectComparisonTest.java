package com.excelninja.domain;

import com.excelninja.domain.exception.HeaderMismatchException;
import com.excelninja.domain.exception.InvalidDocumentStructureException;
import com.excelninja.domain.model.*;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.*;

import static org.apache.logging.log4j.util.Strings.repeat;
import static org.assertj.core.api.Assertions.*;

@DisplayName("값 객체 도입 효과 비교 테스트")
class ValueObjectComparisonTest {

    @Test
    @DisplayName("Before: 원시 타입 사용 - 잘못된 사용 예시")
    void beforeValueObjects_PrimitiveUsage() {
        String emptySheetName = "";
        assertThat(emptySheetName).isEmpty();

        List<String> duplicateHeaders = Arrays.asList("Name", "Name", "Age");
        assertThat(duplicateHeaders).hasSize(3);

        List<List<Object>> inconsistentRows = Arrays.asList(
                Arrays.asList("Hyunsoo", 30, "hyunsoo@example.com"),
                Arrays.asList("Jane", 25),
                Collections.singletonList("Bob")
        );
        assertThat(inconsistentRows.get(0)).hasSize(3);
        assertThat(inconsistentRows.get(1)).hasSize(2);
        assertThat(inconsistentRows.get(2)).hasSize(1);

        assertThatThrownBy(() -> {
            validateSheetName(emptySheetName);
            validateHeaders(duplicateHeaders);
            validateRowConsistency(inconsistentRows, 3);
        });
    }

    private void validateSheetName(String sheetName) {
        if (sheetName == null || sheetName.trim().isEmpty()) {
            throw new RuntimeException("Empty sheet name");
        }
        if (sheetName.contains("/") || sheetName.contains("\\")) {
            throw new RuntimeException("Invalid characters in sheet name");
        }
        if (sheetName.length() > 31) {
            throw new RuntimeException("Sheet name too long");
        }
    }

    private void validateHeaders(List<String> headers) {
        if (headers == null || headers.isEmpty()) {
            throw new RuntimeException("Empty headers");
        }
        Set<String> uniqueHeaders = new HashSet<>();
        for (String header : headers) {
            if (!uniqueHeaders.add(header)) {
                throw new RuntimeException("Duplicate header: " + header);
            }
        }
    }

    private void validateRowConsistency(
            List<List<Object>> rows,
            int expectedColumnCount
    ) {
        for (int i = 0; i < rows.size(); i++) {
            List<Object> row = rows.get(i);
            if (row.size() != expectedColumnCount) {
                throw new RuntimeException("Row " + i + " has " + row.size() + " columns, expected " + expectedColumnCount);
            }
        }
    }

    @Test
    @DisplayName("Before: 원시 타입 사용 시 발생 가능한 런타임 에러들")
    void beforeValueObjects_RuntimeErrors() {

        assertThatThrownBy(() -> validateSheetName(""))
                .isInstanceOf(RuntimeException.class)
                .hasMessageContaining("Empty sheet name");

        assertThatThrownBy(() -> validateHeaders(Arrays.asList("Name", "Name")))
                .isInstanceOf(RuntimeException.class)
                .hasMessageContaining("Duplicate header");

        assertThatThrownBy(() -> validateRowConsistency(Arrays.asList(
                Arrays.asList("Hyunsoo", 30),
                Collections.singletonList("Jane")
        ), 2))
                .isInstanceOf(RuntimeException.class)
                .hasMessageContaining("has 1 columns, expected 2");
    }

    @Test
    @DisplayName("After: 값 객체 사용 - 타입 안전성과 도메인 규칙 보장")
    void afterValueObjects_TypeSafetyAndDomainRules() {
        assertThatThrownBy(() -> new SheetName("")).isInstanceOf(InvalidDocumentStructureException.class);
        assertThatThrownBy(() -> Headers.of("Name", "Name")).isInstanceOf(HeaderMismatchException.class);
        assertThatThrownBy(() -> DocumentRows.of(Arrays.asList(Arrays.asList("Hyunsoo", 30), Collections.singletonList("Jane")), 2)).isInstanceOf(InvalidDocumentStructureException.class);
        assertThatCode(() -> {
            SheetName sheetName = new SheetName("ValidSheet");
            Headers headers = Headers.of("Name", "Age", "Email");
            DocumentRows rows = DocumentRows.of(Arrays.asList(
                    Arrays.asList("Hyunsoo", 30, "hyunsoo@example.com"),
                    Arrays.asList("Jane", 25, "jane@example.com")
            ), 3);

            ExcelDocument document = ExcelDocument.readBuilder()
                    .sheet(sheetName)
                    .headers(headers)
                    .rows(rows)
                    .build();

            assertThat(document.getCellValue(0, "Name")).isEqualTo("Hyunsoo");
            assertThat(document.getCellValue(1, "Age", Integer.class)).isEqualTo(25);

        }).doesNotThrowAnyException();
    }

    @Test
    @DisplayName("Before vs After: 코드 가독성과 안전성 비교")
    void codeReadabilityAndSafetyComparison() {
        String sheetName = "Sheet1";
        List<String> headers = Arrays.asList("Name", "Age");
        List<List<Object>> rows = Collections.singletonList(Arrays.asList("Hyunsoo", 30));

        SheetName meaningfulSheetName = new SheetName("UserData");
        Headers meaningfulHeaders = Headers.of("Name", "Age", "Email");
        DocumentRows meaningfulRows = DocumentRows.of(Collections.singletonList(Arrays.asList("Hyunsoo", 30, "hyunsoo@example.com")), 3);

        assertThat(meaningfulSheetName.isDefault()).isFalse();
        assertThat(meaningfulHeaders.containsHeader("Name")).isTrue();
        assertThat(meaningfulHeaders.getPositionOf("Age")).isEqualTo(1);

        DocumentRow firstRow = meaningfulRows.getRow(0);
        assertThat(firstRow.getValueByHeader(meaningfulHeaders, "Name")).isEqualTo("Hyunsoo");
        assertThat(firstRow.getValueByHeader(meaningfulHeaders, "Age", Integer.class)).isEqualTo(30);
        assertThat(firstRow.isEmpty()).isFalse();
        assertThat(firstRow.hasValue(0)).isTrue();
    }

    @Test
    @DisplayName("값 객체의 표현력 개선 효과")
    void valueObjectExpressiveness() {
        SheetName sheetName = new SheetName("UserData");
        Headers headers = Headers.of("ID", "Name", "Age", "Email");
        DocumentRow row = DocumentRow.of(Arrays.asList(1L, "Hyunsoo", 30, "hyunsoo@example.com"), 1);

        assertThat(sheetName.isDefault()).isFalse();
        assertThat(headers.containsHeader("Name")).isTrue();
        assertThat(headers.getPositionOf("Age")).isEqualTo(2);
        assertThat(row.getValueByHeader(headers, "Name")).isEqualTo("Hyunsoo");
        assertThat(row.getValueByHeader(headers, "Age", Integer.class)).isEqualTo(30);
        assertThat(row.isEmpty()).isFalse();
        assertThat(row.hasValue(1)).isTrue();

        assertThat(row.getRowNumber()).isEqualTo(1);
        assertThat(row.getColumnCount()).isEqualTo(4);
    }

    @Test
    @DisplayName("값 객체의 불변성 vs 가변성 비교")
    void immutabilityComparison() {
        List<String> mutableHeaders = Arrays.asList("Name", "Age");

        Headers originalHeaders = Headers.of("Name", "Age");

        Headers extendedHeaders = originalHeaders.withAdditionalHeader("Email");

        assertThat(originalHeaders.size()).isEqualTo(2);
        assertThat(originalHeaders.size()).isEqualTo(2);
        assertThat(extendedHeaders.size()).isEqualTo(3);

        assertThat(originalHeaders).isSameAs(originalHeaders);
        assertThat(extendedHeaders).isNotSameAs(originalHeaders);
    }

    @Test
    @DisplayName("값 객체의 동등성 비교")
    void valueObjectEquality() {
        SheetName sheet1 = new SheetName("TestSheet");
        SheetName sheet2 = new SheetName("TestSheet");
        SheetName sheet3 = new SheetName("DifferentSheet");

        assertThat(sheet1).isEqualTo(sheet2);
        assertThat(sheet1.hashCode()).isEqualTo(sheet2.hashCode());

        assertThat(sheet1).isNotEqualTo(sheet3);
        assertThat(sheet1.hashCode()).isNotEqualTo(sheet3.hashCode());

        Headers headers1 = Headers.of("Name", "Age");
        Headers headers2 = Headers.of("Name", "Age");
        Headers headers3 = Headers.of("Name", "Email");

        assertThat(headers1).isEqualTo(headers2);
        assertThat(headers1).isNotEqualTo(headers3);
    }

    @Test
    @DisplayName("값 객체의 자기 검증 (Self-Validation)")
    void selfValidation() {
        assertThatThrownBy(() -> new SheetName(null)).isInstanceOf(InvalidDocumentStructureException.class);

        assertThatThrownBy(() -> new SheetName("Sheet/With/Invalid/Chars")).isInstanceOf(InvalidDocumentStructureException.class);

        assertThatThrownBy(() -> new SheetName(repeat("A", 32))).isInstanceOf(InvalidDocumentStructureException.class);

        assertThatThrownBy(() -> new Header("", 0)).isInstanceOf(InvalidDocumentStructureException.class);

        assertThatThrownBy(() -> new Header("ValidName", -1)).isInstanceOf(InvalidDocumentStructureException.class);

        assertThatCode(() -> {
            SheetName validSheet = new SheetName("ValidSheet");
            Header validHeader = new Header("ValidHeader", 0);
            Headers validHeaders = Headers.of("Name", "Age");

            assertThat(validSheet.getValue()).isNotEmpty();
            assertThat(validHeader.getName()).isNotEmpty();
            assertThat(validHeader.getPosition()).isGreaterThanOrEqualTo(0);
            assertThat(validHeaders.size()).isGreaterThan(0);

        }).doesNotThrowAnyException();
    }

    @Test
    @DisplayName("도메인 지식의 응집도 개선")
    void domainKnowledgeCohesion() {

        SheetName sheetName = new SheetName("UserData");
        assertThat(sheetName.isDefault()).isFalse();
        assertThat(SheetName.defaultName().isDefault()).isTrue();

        Headers headers = Headers.of("ID", "Name", "Age", "Email");
        assertThat(headers.containsHeader("Name")).isTrue();
        assertThat(headers.getPositionOf("Age")).isEqualTo(2);

        DocumentRow row = DocumentRow.of(Arrays.asList(1L, "Hyunsoo", 30, "hyunsoo@example.com"), 1);
        assertThat(row.getValueByHeader(headers, "Name")).isEqualTo("Hyunsoo");
        assertThat(row.getValueByHeader(headers, "Age", Integer.class)).isEqualTo(30);
        assertThat(row.isEmpty()).isFalse();
        assertThat(row.hasValue(1)).isTrue();

        assertThat(row.getRowNumber()).isEqualTo(1);
        assertThat(row.getColumnCount()).isEqualTo(4);
    }
}