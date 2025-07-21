package com.excelninja.domain;

import com.excelninja.domain.exception.HeaderMismatchException;
import com.excelninja.domain.exception.InvalidDocumentStructureException;
import com.excelninja.domain.model.*;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

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

    @Test
    @DisplayName("Before: BigDecimal과 LocalDate 처리의 복잡성")
    void beforeValueObjects_ComplexTypeHandling() {
        List<Object> financialRow = Arrays.asList(1001L, "150000.50", "2020-01-15");

        Long accountId = (Long) financialRow.get(0);
        String balanceStr = (String) financialRow.get(1);
        String dateStr = (String) financialRow.get(2);

        BigDecimal balance = new BigDecimal(balanceStr);
        LocalDate openingDate = LocalDate.parse(dateStr);

        assertThat(accountId).isEqualTo(1001L);
        assertThat(balance).isEqualTo(new BigDecimal("150000.50"));
        assertThat(openingDate).isEqualTo(LocalDate.of(2020, 1, 15));

        assertThatThrownBy(() -> {
            List<Object> invalidRow = Arrays.asList(1002L, "invalid_number", "invalid_date");
            new BigDecimal((String) invalidRow.get(1));
        }).isInstanceOf(NumberFormatException.class);
    }

    @Test
    @DisplayName("Before: LocalDateTime 처리 시 발생하는 문제점")
    void beforeValueObjects_LocalDateTimeComplexity() {
        List<Object> eventRow = Arrays.asList(1L, "LOGIN", "2024-01-15T09:30:00", "user123");

        Long eventId = (Long) eventRow.get(0);
        String eventType = (String) eventRow.get(1);
        String timestampStr = (String) eventRow.get(2);
        String userId = (String) eventRow.get(3);

        LocalDateTime timestamp = LocalDateTime.parse(timestampStr);

        assertThat(eventId).isEqualTo(1L);
        assertThat(eventType).isEqualTo("LOGIN");
        assertThat(timestamp).isEqualTo(LocalDateTime.of(2024, 1, 15, 9, 30, 0));
        assertThat(userId).isEqualTo("user123");

        assertThatThrownBy(() -> {
            List<Object> invalidRow = Arrays.asList(2L, "LOGOUT", "invalid_timestamp", "user456");
            LocalDateTime.parse((String) invalidRow.get(2));
        }).isInstanceOf(Exception.class);
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
    @DisplayName("After: 값 객체 사용 - 타입 안전성과 도메인 규칙 보장")
    void afterValueObjects_TypeSafetyAndDomainRules() {
        assertThatThrownBy(() -> new SheetName("")).isInstanceOf(InvalidDocumentStructureException.class);
        assertThatThrownBy(() -> Headers.of("Name", "Name")).isInstanceOf(HeaderMismatchException.class);
        assertThatThrownBy(() -> DocumentRows.of(Arrays.asList(Arrays.asList("Hyunsoo", 30), Collections.singletonList("Jane")), 2)).isInstanceOf(InvalidDocumentStructureException.class);

        assertThatCode(() -> {
            ExcelSheet sheet = ExcelSheet.builder()
                    .name("ValidSheet")
                    .headers("Name", "Age", "Email")
                    .rows(Arrays.asList(
                            Arrays.asList("Hyunsoo", 30, "hyunsoo@example.com"),
                            Arrays.asList("Jane", 25, "jane@example.com")
                    ))
                    .build();

            assertThat(sheet.getCellValue(0, "Name")).isEqualTo("Hyunsoo");
            assertThat(sheet.getCellValue(1, "Age", Integer.class)).isEqualTo(25);

        }).doesNotThrowAnyException();
    }

    @Test
    @DisplayName("After: BigDecimal과 LocalDate를 포함한 안전한 처리")
    void afterValueObjects_SafeBigDecimalAndLocalDateHandling() {
        assertThatCode(() -> {
            ExcelSheet sheet = ExcelSheet.builder()
                    .name("FinancialData")
                    .headers("Account ID", "Balance", "Opening Date")
                    .rows(Arrays.asList(
                            Arrays.asList(1001L, new BigDecimal("150000.50"), LocalDate.of(2020, 1, 15)),
                            Arrays.asList(1002L, new BigDecimal("85000.75"), LocalDate.of(2021, 6, 10))
                    ))
                    .build();

            assertThat(sheet.getCellValue(0, "Balance", BigDecimal.class)).isEqualTo(new BigDecimal("150000.50"));
            assertThat(sheet.getCellValue(1, "Opening Date", LocalDate.class)).isEqualTo(LocalDate.of(2021, 6, 10));

        }).doesNotThrowAnyException();
    }

    @Test
    @DisplayName("After: LocalDateTime을 포함한 안전한 처리")
    void afterValueObjects_SafeLocalDateTimeHandling() {
        assertThatCode(() -> {
            ExcelSheet sheet = ExcelSheet.builder()
                    .name("EventLog")
                    .headers("Event ID", "Event Type", "Timestamp", "User ID")
                    .rows(Arrays.asList(
                            Arrays.asList(1L, "LOGIN", LocalDateTime.of(2024, 1, 15, 9, 30, 0), "user123"),
                            Arrays.asList(2L, "LOGOUT", LocalDateTime.of(2024, 1, 15, 17, 45, 30), "user456")
                    ))
                    .build();

            assertThat(sheet.getCellValue(0, "Timestamp", LocalDateTime.class)).isEqualTo(LocalDateTime.of(2024, 1, 15, 9, 30, 0));
            assertThat(sheet.getCellValue(1, "User ID")).isEqualTo("user456");

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
    @DisplayName("BigDecimal과 LocalDateTime을 포함한 값 객체 표현력")
    void advancedValueObjectExpressiveness() {
        SheetName sheetName = new SheetName("FinancialProducts");
        Headers headers = Headers.of("Product ID", "Name", "Price", "Launch Date", "Last Updated");
        DocumentRow row = DocumentRow.of(Arrays.asList(
                1L,
                "Premium Account",
                new BigDecimal("999.99"),
                LocalDate.of(2024, 1, 1),
                LocalDateTime.of(2024, 1, 15, 10, 30, 0)
        ), 1);

        assertThat(sheetName.getValue()).isEqualTo("FinancialProducts");
        assertThat(headers.containsHeader("Price")).isTrue();
        assertThat(headers.getPositionOf("Launch Date")).isEqualTo(3);

        assertThat(row.getValueByHeader(headers, "Price", BigDecimal.class)).isEqualTo(new BigDecimal("999.99"));
        assertThat(row.getValueByHeader(headers, "Launch Date", LocalDate.class)).isEqualTo(LocalDate.of(2024, 1, 1));
        assertThat(row.getValueByHeader(headers, "Last Updated", LocalDateTime.class)).isEqualTo(LocalDateTime.of(2024, 1, 15, 10, 30, 0));

        assertThat(row.getRowNumber()).isEqualTo(1);
        assertThat(row.getColumnCount()).isEqualTo(5);
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
    @DisplayName("복잡한 타입을 포함한 DocumentRow 동등성")
    void complexTypeDocumentRowEquality() {
        DocumentRow row1 = DocumentRow.of(Arrays.asList(
                1L,
                "Product A",
                new BigDecimal("99.99"),
                LocalDate.of(2024, 1, 1),
                LocalDateTime.of(2024, 1, 15, 10, 30, 0)
        ), 1);

        DocumentRow row2 = DocumentRow.of(Arrays.asList(
                1L,
                "Product A",
                new BigDecimal("99.99"),
                LocalDate.of(2024, 1, 1),
                LocalDateTime.of(2024, 1, 15, 10, 30, 0)
        ), 1);

        DocumentRow row3 = DocumentRow.of(Arrays.asList(
                2L,
                "Product B",
                new BigDecimal("149.99"),
                LocalDate.of(2024, 2, 1),
                LocalDateTime.of(2024, 2, 10, 14, 45, 30)
        ), 2);

        assertThat(row1).isEqualTo(row2);
        assertThat(row1.hashCode()).isEqualTo(row2.hashCode());
        assertThat(row1).isNotEqualTo(row3);
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

    @Test
    @DisplayName("복잡한 타입을 포함한 도메인 지식 응집도")
    void complexTypeDomainKnowledgeCohesion() {
        SheetName sheetName = new SheetName("FinancialData");
        Headers headers = Headers.of("Account ID", "Balance", "Rate", "Opening Date", "Last Updated");
        DocumentRow row = DocumentRow.of(Arrays.asList(
                1001L,
                new BigDecimal("150000.50"),
                new BigDecimal("3.25"),
                LocalDate.of(2020, 1, 15),
                LocalDateTime.of(2024, 1, 15, 10, 30, 0)
        ), 1);

        assertThat(sheetName.getValue()).isEqualTo("FinancialData");
        assertThat(headers.containsHeader("Balance")).isTrue();
        assertThat(headers.getPositionOf("Opening Date")).isEqualTo(3);

        BigDecimal balance = row.getValueByHeader(headers, "Balance", BigDecimal.class);
        LocalDate openingDate = row.getValueByHeader(headers, "Opening Date", LocalDate.class);
        LocalDateTime lastUpdated = row.getValueByHeader(headers, "Last Updated", LocalDateTime.class);

        assertThat(balance).isEqualTo(new BigDecimal("150000.50"));
        assertThat(openingDate).isEqualTo(LocalDate.of(2020, 1, 15));
        assertThat(lastUpdated).isEqualTo(LocalDateTime.of(2024, 1, 15, 10, 30, 0));

        assertThat(row.isEmpty()).isFalse();
        assertThat(row.hasValue(2)).isTrue();
        assertThat(row.getRowNumber()).isEqualTo(1);
        assertThat(row.getColumnCount()).isEqualTo(5);
    }

    private static String repeat(final String str, final int count) {
        if (str == null) {
            throw new NullPointerException("str");
        }
        if (count < 0) {
            throw new IllegalArgumentException("count");
        }
        StringBuilder sb = new StringBuilder();
        for (int index = 0; index < count; ++index) {
            sb.append(str);
        }
        return sb.toString();
    }
}