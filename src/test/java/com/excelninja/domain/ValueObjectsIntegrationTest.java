package com.excelninja.domain;

import com.excelninja.application.facade.NinjaExcel;
import com.excelninja.domain.annotation.ExcelReadColumn;
import com.excelninja.domain.annotation.ExcelWriteColumn;
import com.excelninja.domain.exception.HeaderMismatchException;
import com.excelninja.domain.exception.InvalidDocumentStructureException;
import com.excelninja.domain.model.*;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

@DisplayName("값 객체 통합 테스트 - 새로운 API")
class ValueObjectsIntegrationTest {

    @TempDir
    Path tempDir;

    @Test
    @DisplayName("값 객체를 사용한 완전한 Excel 처리 플로우")
    void completeExcelProcessingWithValueObjects() throws IOException {
        List<UserDto> users = Arrays.asList(
                new UserDto("Hyunsoo", 30, "hyunsoo@example.com", true),
                new UserDto("Eunmi", 25, "eunmi@example.com", false),
                new UserDto("naruto", 35, "naruto@example.com", true)
        );

        ExcelWorkbook workbook = ExcelWorkbook.builder()
                .sheet("UserData", users)
                .build();

        Path testFile = tempDir.resolve("users.xlsx");
        NinjaExcel.write(workbook, testFile.toFile());

        List<UserDto> readUsers = NinjaExcel.read(testFile.toFile(), UserDto.class);

        assertThat(readUsers).hasSize(3);

        UserDto firstUser = readUsers.get(0);
        assertThat(firstUser.getName()).isEqualTo("Hyunsoo");
        assertThat(firstUser.getAge()).isEqualTo(30);
        assertThat(firstUser.getEmail()).isEqualTo("hyunsoo@example.com");
        assertThat(firstUser.isActive()).isTrue();

        UserDto secondUser = readUsers.get(1);
        assertThat(secondUser.getName()).isEqualTo("Eunmi");
        assertThat(secondUser.getAge()).isEqualTo(25);
        assertThat(secondUser.isActive()).isFalse();
    }

    @Test
    @DisplayName("BigDecimal과 LocalDate를 사용한 금융 데이터 처리")
    void financialDataProcessingWithBigDecimalAndLocalDate() throws IOException {
        List<AccountDto> accounts = Arrays.asList(
                new AccountDto(1001L, new BigDecimal("150000.50"), new BigDecimal("3.25"), LocalDate.of(2020, 1, 15)),
                new AccountDto(1002L, new BigDecimal("85000.75"), new BigDecimal("2.80"), LocalDate.of(2021, 6, 10)),
                new AccountDto(1003L, new BigDecimal("250000.00"), new BigDecimal("4.15"), LocalDate.of(2019, 11, 5))
        );

        ExcelWorkbook workbook = ExcelWorkbook.builder()
                .sheet("FinancialData", accounts)
                .build();

        Path testFile = tempDir.resolve("financial.xlsx");
        NinjaExcel.write(workbook, testFile.toFile());

        List<AccountDto> readAccounts = NinjaExcel.read(testFile.toFile(), AccountDto.class);

        assertThat(readAccounts).hasSize(3);

        AccountDto firstAccount = readAccounts.get(0);
        assertThat(firstAccount.getAccountId()).isEqualTo(1001L);
        assertThat(firstAccount.getBalance()).isEqualTo(new BigDecimal("150000.5"));
        assertThat(firstAccount.getInterestRate()).isEqualTo(new BigDecimal("3.25"));
        assertThat(firstAccount.getOpeningDate()).isEqualTo(LocalDate.of(2020, 1, 15));
    }

    @Test
    @DisplayName("LocalDateTime을 사용한 이벤트 로그 처리")
    void eventLogProcessingWithLocalDateTime() throws IOException {
        List<EventLogDto> events = Arrays.asList(
                new EventLogDto(1L, "LOGIN", LocalDateTime.of(2024, 1, 15, 9, 30, 0), "user123"),
                new EventLogDto(2L, "LOGOUT", LocalDateTime.of(2024, 1, 15, 17, 45, 30), "user123"),
                new EventLogDto(3L, "PURCHASE", LocalDateTime.of(2024, 1, 16, 14, 20, 15), "user456")
        );

        ExcelWorkbook workbook = ExcelWorkbook.builder()
                .sheet("EventLog", events)
                .build();

        Path testFile = tempDir.resolve("events.xlsx");
        NinjaExcel.write(workbook, testFile.toFile());

        List<EventLogDto> readEvents = NinjaExcel.read(testFile.toFile(), EventLogDto.class);

        assertThat(readEvents).hasSize(3);

        EventLogDto firstEvent = readEvents.get(0);
        assertThat(firstEvent.getEventId()).isEqualTo(1L);
        assertThat(firstEvent.getEventType()).isEqualTo("LOGIN");
        assertThat(firstEvent.getTimestamp()).isEqualTo(LocalDateTime.of(2024, 1, 15, 9, 30, 0));
        assertThat(firstEvent.getUserId()).isEqualTo("user123");
    }

    @Test
    @DisplayName("모든 지원 타입을 포함한 복합 데이터 처리")
    void comprehensiveDataProcessingWithAllSupportedTypes() throws IOException {
        List<ComprehensiveDto> products = Arrays.asList(
                new ComprehensiveDto(1L, "Product A", new BigDecimal("99.99"), LocalDate.of(2024, 1, 1), LocalDateTime.of(2024, 1, 15, 10, 30, 0), true),
                new ComprehensiveDto(2L, "Product B", new BigDecimal("149.50"), LocalDate.of(2024, 2, 1), LocalDateTime.of(2024, 2, 10, 14, 45, 30), false)
        );

        ExcelWorkbook workbook = ExcelWorkbook.builder()
                .sheet("ComprehensiveData", products)
                .build();

        Path testFile = tempDir.resolve("comprehensive.xlsx");
        NinjaExcel.write(workbook, testFile.toFile());

        List<ComprehensiveDto> readProducts = NinjaExcel.read(testFile.toFile(), ComprehensiveDto.class);

        assertThat(readProducts).hasSize(2);

        ComprehensiveDto firstProduct = readProducts.get(0);
        assertThat(firstProduct.getId()).isEqualTo(1L);
        assertThat(firstProduct.getName()).isEqualTo("Product A");
        assertThat(firstProduct.getPrice()).isEqualTo(new BigDecimal("99.99"));
        assertThat(firstProduct.getCreatedDate()).isEqualTo(LocalDate.of(2024, 1, 1));
        assertThat(firstProduct.getLastModified()).isEqualTo(LocalDateTime.of(2024, 1, 15, 10, 30, 0));
        assertThat(firstProduct.isActive()).isTrue();
    }

    @Test
    @DisplayName("다중 시트 처리")
    void multiSheetProcessing() throws IOException {
        List<UserDto> users = Arrays.asList(
                new UserDto("Alice", 28, "alice@example.com", true),
                new UserDto("Charlie", 32, "charlie@example.com", false)
        );

        List<AccountDto> accounts = Arrays.asList(
                new AccountDto(1001L, new BigDecimal("50000"), new BigDecimal("2.5"), LocalDate.of(2023, 1, 1))
        );

        ExcelWorkbook workbook = ExcelWorkbook.builder()
                .sheet("Users", users)
                .sheet("Accounts", accounts)
                .metadata(new WorkbookMetadata("Test Author", "Multi-Sheet Test", LocalDateTime.now()))
                .build();

        Path testFile = tempDir.resolve("multi-sheet.xlsx");
        NinjaExcel.write(workbook, testFile.toFile());

        List<UserDto> readUsers = NinjaExcel.readSheet(testFile.toFile(), "Users", UserDto.class);
        List<AccountDto> readAccounts = NinjaExcel.readSheet(testFile.toFile(), "Accounts", AccountDto.class);

        assertThat(readUsers).hasSize(2);
        assertThat(readAccounts).hasSize(1);

        assertThat(readUsers.get(0).getName()).isEqualTo("Alice");
        assertThat(readAccounts.get(0).getAccountId()).isEqualTo(1001L);

        List<String> sheetNames = NinjaExcel.getSheetNames(testFile.toFile());
        assertThat(sheetNames).containsExactly("Users", "Accounts");
    }

    @Test
    @DisplayName("커스텀 시트 빌더를 사용한 데이터 처리")
    void customSheetBuilderProcessing() throws IOException {
        ExcelSheet customSheet = ExcelSheet.builder()
                .name("CustomData")
                .headers("ID", "Name", "Price", "Date")
                .rows(Arrays.asList(
                        Arrays.asList(1L, "Item1", new BigDecimal("99.99"), LocalDate.of(2024, 1, 15)),
                        Arrays.asList(2L, "Item2", new BigDecimal("149.50"), LocalDate.of(2024, 2, 20))
                ))
                .columnWidth(0, 3000)
                .columnWidth(1, 5000)
                .build();

        ExcelWorkbook workbook = ExcelWorkbook.builder()
                .sheet("Custom", customSheet)
                .build();

        Path testFile = tempDir.resolve("custom.xlsx");
        NinjaExcel.write(workbook, testFile.toFile());

        assertThat(testFile.toFile()).exists();
        assertThat(testFile.toFile().length()).isGreaterThan(0);
    }

    @Test
    @DisplayName("ExcelSheet 빌더에서 행/열 불일치 시 InvalidDocumentStructureException 발생")
    void rowColumnMismatch() {
        assertThatThrownBy(() ->
                ExcelSheet.builder()
                        .name("TestSheet")
                        .headers("Header1", "Header2")
                        .rows(Arrays.asList(
                                Arrays.asList("Value1")
                        ))
                        .build()
        )
                .isInstanceOf(InvalidDocumentStructureException.class)
                .hasMessageContaining("has 1 columns but expected 2");
    }

    @Test
    @DisplayName("중복 헤더명으로 시트 생성 시 HeaderMismatchException 발생")
    void duplicateHeadersInSheet() {
        assertThatThrownBy(() ->
                ExcelSheet.builder()
                        .name("TestSheet")
                        .headers("Header1", "Header1")
                        .build()
        )
                .isInstanceOf(HeaderMismatchException.class)
                .hasMessageContaining("duplicate");
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

        ExcelSheet sheet = ExcelSheet.builder()
                .name(sheetName.getValue())
                .headers(headers.getHeaderNames())
                .rows(Arrays.asList(
                        Arrays.asList("A", "B", "C"),
                        Arrays.asList("X", "Y", "Z")
                ))
                .build();

        Headers retrievedHeaders = sheet.getHeaders();
        DocumentRows retrievedRows = sheet.getRows();

        assertThatThrownBy(() -> retrievedHeaders.getHeaders().add(new Header("NewHeader", 3)))
                .isInstanceOf(UnsupportedOperationException.class);

        assertThatThrownBy(() -> retrievedRows.getRows().add(DocumentRow.of(Arrays.asList("New", "Row", "Data"), 3)))
                .isInstanceOf(UnsupportedOperationException.class);

        SheetName originalSheetName = sheet.getName();
        assertThat(originalSheetName.getValue()).isEqualTo("ImmutableTest");

        SheetName newSheetName = new SheetName("NewSheet");
        assertThat(originalSheetName.getValue()).isEqualTo("ImmutableTest");
        assertThat(newSheetName.getValue()).isEqualTo("NewSheet");
    }

    @Test
    @DisplayName("전체 시트 읽기 테스트 - 같은 타입")
    void readAllSheetsTest() throws IOException {
        List<UserDto> engineeringUsers = Arrays.asList(
                new UserDto("Alice", 28, "alice@example.com", true)
        );

        List<UserDto> marketingUsers = Arrays.asList(
                new UserDto("Bob", 32, "bob@example.com", false)
        );

        ExcelWorkbook workbook = ExcelWorkbook.builder()
                .sheet("Engineering", engineeringUsers)
                .sheet("Marketing", marketingUsers)
                .build();

        Path testFile = tempDir.resolve("all-user-sheets.xlsx");
        NinjaExcel.write(workbook, testFile.toFile());

        Map<String, List<UserDto>> allUserSheets = NinjaExcel.readAllSheets(testFile.toFile(), UserDto.class);

        assertThat(allUserSheets).hasSize(2);
        assertThat(allUserSheets.get("Engineering")).hasSize(1);
        assertThat(allUserSheets.get("Marketing")).hasSize(1);
        assertThat(allUserSheets.get("Engineering").get(0).getName()).isEqualTo("Alice");
        assertThat(allUserSheets.get("Marketing").get(0).getName()).isEqualTo("Bob");
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

        public ComprehensiveDto(
                Long id,
                String name,
                BigDecimal price,
                LocalDate createdDate,
                LocalDateTime lastModified,
                Boolean isActive
        ) {
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