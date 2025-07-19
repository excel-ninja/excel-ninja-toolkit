package com.excelninja.domain;

import com.excelninja.domain.annotation.ExcelReadColumn;
import com.excelninja.domain.annotation.ExcelWriteColumn;
import com.excelninja.domain.exception.HeaderMismatchException;
import com.excelninja.domain.exception.InvalidDocumentStructureException;
import com.excelninja.domain.model.*;
import com.excelninja.infrastructure.converter.DefaultConverter;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
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

        ExcelDocument document = ExcelDocument.reader()
                .sheet(sheetName)
                .headers(headers)
                .rows(rows)
                .create();

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
    @DisplayName("BigDecimal과 LocalDate를 사용한 금융 데이터 처리")
    void financialDataProcessingWithBigDecimalAndLocalDate() {
        SheetName sheetName = new SheetName("FinancialData");
        Headers headers = Headers.of("Account ID", "Balance", "Interest Rate", "Opening Date");
        DocumentRows rows = DocumentRows.of(Arrays.asList(
                Arrays.asList(1001L, new BigDecimal("150000.50"), new BigDecimal("3.25"), LocalDate.of(2020, 1, 15)),
                Arrays.asList(1002L, new BigDecimal("85000.75"), new BigDecimal("2.80"), LocalDate.of(2021, 6, 10)),
                Arrays.asList(1003L, new BigDecimal("250000.00"), new BigDecimal("4.15"), LocalDate.of(2019, 11, 5))
        ), 4);

        ExcelDocument document = ExcelDocument.reader()
                .sheet(sheetName)
                .headers(headers)
                .rows(rows)
                .create();

        List<AccountDto> accounts = document.convertToEntities(AccountDto.class, new DefaultConverter());

        assertThat(accounts).hasSize(3);

        AccountDto firstAccount = accounts.get(0);
        assertThat(firstAccount.getAccountId()).isEqualTo(1001L);
        assertThat(firstAccount.getBalance()).isEqualTo(new BigDecimal("150000.50"));
        assertThat(firstAccount.getInterestRate()).isEqualTo(new BigDecimal("3.25"));
        assertThat(firstAccount.getOpeningDate()).isEqualTo(LocalDate.of(2020, 1, 15));

        assertThat(document.getCellValue(0, "Balance", BigDecimal.class)).isEqualTo(new BigDecimal("150000.50"));
        assertThat(document.getCellValue(1, "Opening Date", LocalDate.class)).isEqualTo(LocalDate.of(2021, 6, 10));
    }

    @Test
    @DisplayName("LocalDateTime을 사용한 이벤트 로그 처리")
    void eventLogProcessingWithLocalDateTime() {
        SheetName sheetName = new SheetName("EventLog");
        Headers headers = Headers.of("Event ID", "Event Type", "Timestamp", "User ID");
        DocumentRows rows = DocumentRows.of(Arrays.asList(
                Arrays.asList(1L, "LOGIN", LocalDateTime.of(2024, 1, 15, 9, 30, 0), "user123"),
                Arrays.asList(2L, "LOGOUT", LocalDateTime.of(2024, 1, 15, 17, 45, 30), "user123"),
                Arrays.asList(3L, "PURCHASE", LocalDateTime.of(2024, 1, 16, 14, 20, 15), "user456")
        ), 4);

        ExcelDocument document = ExcelDocument.reader()
                .sheet(sheetName)
                .headers(headers)
                .rows(rows)
                .create();

        List<EventLogDto> events = document.convertToEntities(EventLogDto.class, new DefaultConverter());

        assertThat(events).hasSize(3);

        EventLogDto firstEvent = events.get(0);
        assertThat(firstEvent.getEventId()).isEqualTo(1L);
        assertThat(firstEvent.getEventType()).isEqualTo("LOGIN");
        assertThat(firstEvent.getTimestamp()).isEqualTo(LocalDateTime.of(2024, 1, 15, 9, 30, 0));
        assertThat(firstEvent.getUserId()).isEqualTo("user123");

        assertThat(document.getCellValue(1, "Timestamp", LocalDateTime.class)).isEqualTo(LocalDateTime.of(2024, 1, 15, 17, 45, 30));
    }

    @Test
    @DisplayName("모든 지원 타입을 포함한 복합 데이터 처리")
    void comprehensiveDataProcessingWithAllSupportedTypes() {
        SheetName sheetName = new SheetName("ComprehensiveData");
        Headers headers = Headers.of("ID", "Name", "Price", "Created Date", "Last Modified", "Is Active");
        DocumentRows rows = DocumentRows.of(Arrays.asList(
                Arrays.asList(1L, "Product A", new BigDecimal("99.99"), LocalDate.of(2024, 1, 1), LocalDateTime.of(2024, 1, 15, 10, 30, 0), true),
                Arrays.asList(2L, "Product B", new BigDecimal("149.50"), LocalDate.of(2024, 2, 1), LocalDateTime.of(2024, 2, 10, 14, 45, 30), false)
        ), 6);

        ExcelDocument document = ExcelDocument.reader()
                .sheet(sheetName)
                .headers(headers)
                .rows(rows)
                .create();

        List<ComprehensiveDto> products = document.convertToEntities(ComprehensiveDto.class, new DefaultConverter());

        assertThat(products).hasSize(2);

        ComprehensiveDto firstProduct = products.get(0);
        assertThat(firstProduct.getId()).isEqualTo(1L);
        assertThat(firstProduct.getName()).isEqualTo("Product A");
        assertThat(firstProduct.getPrice()).isEqualTo(new BigDecimal("99.99"));
        assertThat(firstProduct.getCreatedDate()).isEqualTo(LocalDate.of(2024, 1, 1));
        assertThat(firstProduct.getLastModified()).isEqualTo(LocalDateTime.of(2024, 1, 15, 10, 30, 0));
        assertThat(firstProduct.isActive()).isTrue();
    }

    @Test
    @DisplayName("값 객체를 사용한 엔티티에서 Excel 문서 생성")
    void createExcelDocumentFromEntitiesWithValueObjects() {
        List<UserDto> users = Arrays.asList(
                new UserDto("Alice", 28, "alice@example.com", true),
                new UserDto("Charlie", 32, "charlie@example.com", false)
        );

        ExcelDocument document = ExcelDocument.writer().objects(users).sheetName("UserReport").create();

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
    @DisplayName("BigDecimal과 LocalDateTime을 포함한 엔티티 생성 테스트")
    void createEntityWithBigDecimalAndLocalDateTime() {
        List<ComprehensiveDto> products = Arrays.asList(
                new ComprehensiveDto(1L, "Premium Product", new BigDecimal("999.99"),
                        LocalDate.of(2024, 1, 1), LocalDateTime.of(2024, 1, 15, 10, 30, 0), true),
                new ComprehensiveDto(2L, "Standard Product", new BigDecimal("49.99"),
                        LocalDate.of(2024, 2, 1), LocalDateTime.of(2024, 2, 10, 14, 45, 30), false)
        );

        ExcelDocument document = ExcelDocument.writer().objects(products).sheetName("ProductCatalog").create();

        assertThat(document.getSheetName().getValue()).isEqualTo("ProductCatalog");
        assertThat(document.getHeaders().size()).isEqualTo(6);
        assertThat(document.getRows().size()).isEqualTo(2);

        DocumentRow firstRow = document.getRows().getRow(0);
        assertThat(firstRow.getValueByHeader(document.getHeaders(), "Price")).isEqualTo(new BigDecimal("999.99"));
        assertThat(firstRow.getValueByHeader(document.getHeaders(), "Created Date")).isEqualTo(LocalDate.of(2024, 1, 1));
        assertThat(firstRow.getValueByHeader(document.getHeaders(), "Last Modified")).isEqualTo(LocalDateTime.of(2024, 1, 15, 10, 30, 0));
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
    @DisplayName("값 객체의 불변성 보장")
    void valueObjectImmutability() {
        SheetName sheetName = new SheetName("ImmutableTest");
        Headers headers = Headers.of("Col1", "Col2", "Col3");
        DocumentRows rows = DocumentRows.of(Arrays.asList(
                Arrays.asList("A", "B", "C"),
                Arrays.asList("X", "Y", "Z")
        ), 3);

        ExcelDocument document = ExcelDocument.reader()
                .sheet(sheetName)
                .headers(headers)
                .rows(rows)
                .create();

        Headers retrievedHeaders = document.getHeaders();
        DocumentRows retrievedRows = document.getRows();

        assertThatThrownBy(() -> retrievedHeaders.getHeaders().add(new Header("NewHeader", 3)))
                .isInstanceOf(UnsupportedOperationException.class);

        assertThatThrownBy(() -> retrievedRows.getRows().add(DocumentRow.of(Arrays.asList("New", "Row", "Data"), 3)))
                .isInstanceOf(UnsupportedOperationException.class);

        SheetName originalSheetName = document.getSheetName();
        assertThat(originalSheetName.getValue()).isEqualTo("ImmutableTest");

        SheetName newSheetName = new SheetName("NewSheet");
        assertThat(originalSheetName.getValue()).isEqualTo("ImmutableTest");
        assertThat(newSheetName.getValue()).isEqualTo("NewSheet");
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

        public UserDto(String name, Integer age, String email, Boolean active) {
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

    public static class AccountDto {
        @ExcelReadColumn(headerName = "Account ID")
        @ExcelWriteColumn(headerName = "Account ID", order = 1)
        private Long accountId;

        @ExcelReadColumn(headerName = "Balance")
        @ExcelWriteColumn(headerName = "Balance", order = 2)
        private BigDecimal balance;

        @ExcelReadColumn(headerName = "Interest Rate")
        @ExcelWriteColumn(headerName = "Interest Rate", order = 3)
        private BigDecimal interestRate;

        @ExcelReadColumn(headerName = "Opening Date")
        @ExcelWriteColumn(headerName = "Opening Date", order = 4)
        private LocalDate openingDate;

        public AccountDto() {}

        public AccountDto(Long accountId, BigDecimal balance, BigDecimal interestRate, LocalDate openingDate) {
            this.accountId = accountId;
            this.balance = balance;
            this.interestRate = interestRate;
            this.openingDate = openingDate;
        }

        public Long getAccountId() {return accountId;}
        public void setAccountId(Long accountId) {this.accountId = accountId;}
        public BigDecimal getBalance() {return balance;}
        public void setBalance(BigDecimal balance) {this.balance = balance;}
        public BigDecimal getInterestRate() {return interestRate;}
        public void setInterestRate(BigDecimal interestRate) {this.interestRate = interestRate;}
        public LocalDate getOpeningDate() {return openingDate;}
        public void setOpeningDate(LocalDate openingDate) {this.openingDate = openingDate;}
    }

    public static class EventLogDto {
        @ExcelReadColumn(headerName = "Event ID")
        @ExcelWriteColumn(headerName = "Event ID", order = 1)
        private Long eventId;

        @ExcelReadColumn(headerName = "Event Type")
        @ExcelWriteColumn(headerName = "Event Type", order = 2)
        private String eventType;

        @ExcelReadColumn(headerName = "Timestamp")
        @ExcelWriteColumn(headerName = "Timestamp", order = 3)
        private LocalDateTime timestamp;

        @ExcelReadColumn(headerName = "User ID")
        @ExcelWriteColumn(headerName = "User ID", order = 4)
        private String userId;

        public EventLogDto() {}

        public EventLogDto(Long eventId, String eventType, LocalDateTime timestamp, String userId) {
            this.eventId = eventId;
            this.eventType = eventType;
            this.timestamp = timestamp;
            this.userId = userId;
        }

        public Long getEventId() {return eventId;}
        public void setEventId(Long eventId) {this.eventId = eventId;}
        public String getEventType() {return eventType;}
        public void setEventType(String eventType) {this.eventType = eventType;}
        public LocalDateTime getTimestamp() {return timestamp;}
        public void setTimestamp(LocalDateTime timestamp) {this.timestamp = timestamp;}
        public String getUserId() {return userId;}
        public void setUserId(String userId) {this.userId = userId;}
    }

    public static class ComprehensiveDto {
        @ExcelReadColumn(headerName = "ID")
        @ExcelWriteColumn(headerName = "ID", order = 1)
        private Long id;

        @ExcelReadColumn(headerName = "Name")
        @ExcelWriteColumn(headerName = "Name", order = 2)
        private String name;

        @ExcelReadColumn(headerName = "Price")
        @ExcelWriteColumn(headerName = "Price", order = 3)
        private BigDecimal price;

        @ExcelReadColumn(headerName = "Created Date")
        @ExcelWriteColumn(headerName = "Created Date", order = 4)
        private LocalDate createdDate;

        @ExcelReadColumn(headerName = "Last Modified")
        @ExcelWriteColumn(headerName = "Last Modified", order = 5)
        private LocalDateTime lastModified;

        @ExcelReadColumn(headerName = "Is Active")
        @ExcelWriteColumn(headerName = "Is Active", order = 6)
        private Boolean isActive;

        public ComprehensiveDto() {}

        public ComprehensiveDto(Long id, String name, BigDecimal price, LocalDate createdDate, LocalDateTime lastModified, Boolean isActive) {
            this.id = id;
            this.name = name;
            this.price = price;
            this.createdDate = createdDate;
            this.lastModified = lastModified;
            this.isActive = isActive;
        }

        public Long getId() {return id;}
        public void setId(Long id) {this.id = id;}
        public String getName() {return name;}
        public void setName(String name) {this.name = name;}
        public BigDecimal getPrice() {return price;}
        public void setPrice(BigDecimal price) {this.price = price;}
        public LocalDate getCreatedDate() {return createdDate;}
        public void setCreatedDate(LocalDate createdDate) {this.createdDate = createdDate;}
        public LocalDateTime getLastModified() {return lastModified;}
        public void setLastModified(LocalDateTime lastModified) {this.lastModified = lastModified;}
        public Boolean isActive() {return isActive;}
        public void setIsActive(Boolean isActive) {this.isActive = isActive;}
    }
}