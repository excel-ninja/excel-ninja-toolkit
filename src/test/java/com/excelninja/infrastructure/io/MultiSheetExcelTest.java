package com.excelninja.infrastructure.io;

import com.excelninja.application.facade.NinjaExcel;
import com.excelninja.domain.annotation.ExcelReadColumn;
import com.excelninja.domain.annotation.ExcelWriteColumn;
import com.excelninja.domain.model.ExcelWorkbook;
import com.excelninja.domain.model.WorkbookMetadata;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

@DisplayName("다중 시트 Excel 테스트")
class MultiSheetExcelTest {

    @TempDir
    Path tempDir;

    public static class UserDto {
        @ExcelReadColumn(headerName = "ID")
        @ExcelWriteColumn(headerName = "ID", order = 1)
        private Long id;

        @ExcelReadColumn(headerName = "Name")
        @ExcelWriteColumn(headerName = "Name", order = 2)
        private String name;

        @ExcelReadColumn(headerName = "Age")
        @ExcelWriteColumn(headerName = "Age", order = 3)
        private Integer age;

        @ExcelReadColumn(headerName = "Email")
        @ExcelWriteColumn(headerName = "Email", order = 4)
        private String email;

        public UserDto() {}

        public UserDto(
                Long id,
                String name,
                Integer age,
                String email
        ) {
            this.id = id;
            this.name = name;
            this.age = age;
            this.email = email;
        }

        public Long getId() {return id;}

        public void setId(Long id) {this.id = id;}

        public String getName() {return name;}

        public void setName(String name) {this.name = name;}

        public Integer getAge() {return age;}

        public void setAge(Integer age) {this.age = age;}

        public String getEmail() {return email;}

        public void setEmail(String email) {this.email = email;}
    }

    public static class DepartmentDto {
        @ExcelReadColumn(headerName = "ID")
        @ExcelWriteColumn(headerName = "ID", order = 1)
        private Long id;

        @ExcelReadColumn(headerName = "Name")
        @ExcelWriteColumn(headerName = "Name", order = 2)
        private String name;

        @ExcelReadColumn(headerName = "Budget")
        @ExcelWriteColumn(headerName = "Budget", order = 3)
        private BigDecimal budget;

        @ExcelReadColumn(headerName = "Head Count")
        @ExcelWriteColumn(headerName = "Head Count", order = 4)
        private Integer headCount;

        public DepartmentDto() {}

        public DepartmentDto(
                Long id,
                String name,
                BigDecimal budget,
                Integer headCount
        ) {
            this.id = id;
            this.name = name;
            this.budget = budget;
            this.headCount = headCount;
        }

        public Long getId() {return id;}

        public void setId(Long id) {this.id = id;}

        public String getName() {return name;}

        public void setName(String name) {this.name = name;}

        public BigDecimal getBudget() {return budget;}

        public void setBudget(BigDecimal budget) {this.budget = budget;}

        public Integer getHeadCount() {return headCount;}

        public void setHeadCount(Integer headCount) {this.headCount = headCount;}
    }

    public static class ProjectDto {
        @ExcelReadColumn(headerName = "ID")
        @ExcelWriteColumn(headerName = "ID", order = 1)
        private Long id;

        @ExcelReadColumn(headerName = "Name")
        @ExcelWriteColumn(headerName = "Name", order = 2)
        private String name;

        @ExcelReadColumn(headerName = "Start Date")
        @ExcelWriteColumn(headerName = "Start Date", order = 3)
        private LocalDate startDate;

        @ExcelReadColumn(headerName = "End Date")
        @ExcelWriteColumn(headerName = "End Date", order = 4)
        private LocalDateTime endDate;

        @ExcelReadColumn(headerName = "Budget")
        @ExcelWriteColumn(headerName = "Budget", order = 5)
        private BigDecimal budget;

        public ProjectDto() {}

        public ProjectDto(
                Long id,
                String name,
                LocalDate startDate,
                LocalDateTime endDate,
                BigDecimal budget
        ) {
            this.id = id;
            this.name = name;
            this.startDate = startDate;
            this.endDate = endDate;
            this.budget = budget;
        }

        public Long getId() {return id;}

        public void setId(Long id) {this.id = id;}

        public String getName() {return name;}

        public void setName(String name) {this.name = name;}

        public LocalDate getStartDate() {return startDate;}

        public void setStartDate(LocalDate startDate) {this.startDate = startDate;}

        public LocalDateTime getEndDate() {return endDate;}

        public void setEndDate(LocalDateTime endDate) {this.endDate = endDate;}

        public BigDecimal getBudget() {return budget;}

        public void setBudget(BigDecimal budget) {this.budget = budget;}
    }

    @Test
    @DisplayName("다중 시트 생성 및 읽기 테스트")
    void createAndReadMultipleSheets() throws Exception {
        List<UserDto> users = Arrays.asList(
                new UserDto(1L, "Alice Johnson", 30, "alice@company.com"),
                new UserDto(2L, "Bob Smith", 25, "bob@company.com"),
                new UserDto(3L, "Charlie Brown", 35, "charlie@company.com")
        );

        List<DepartmentDto> departments = Arrays.asList(
                new DepartmentDto(1L, "Engineering", new BigDecimal("500000"), 15),
                new DepartmentDto(2L, "Marketing", new BigDecimal("300000"), 8),
                new DepartmentDto(3L, "Sales", new BigDecimal("400000"), 12)
        );

        List<ProjectDto> projects = Arrays.asList(
                new ProjectDto(1L, "Project Alpha", LocalDate.of(2024, 1, 15),
                        LocalDateTime.of(2024, 12, 31, 23, 59, 59), new BigDecimal("150000")),
                new ProjectDto(2L, "Project Beta", LocalDate.of(2024, 3, 1),
                        LocalDateTime.of(2025, 6, 30, 23, 59, 59), new BigDecimal("200000"))
        );

        ExcelWorkbook workbook = ExcelWorkbook.builder()
                .sheet("Users", users)
                .sheet("Departments", departments)
                .sheet("Projects", projects)
                .metadata(new WorkbookMetadata("Test Author", "Multi-Sheet Company Report", LocalDateTime.now()))
                .build();

        Path testFile = tempDir.resolve("multi_sheet_company.xlsx");
        NinjaExcel.write(workbook, testFile.toFile());

        assertThat(testFile.toFile()).exists();
        assertThat(testFile.toFile().length()).isGreaterThan(0);

        List<String> sheetNames = NinjaExcel.getSheetNames(testFile.toFile());
        assertThat(sheetNames).containsExactly("Users", "Departments", "Projects");

        List<UserDto> readUsers = NinjaExcel.readSheet(testFile.toFile(), "Users", UserDto.class);
        List<DepartmentDto> readDepartments = NinjaExcel.readSheet(testFile.toFile(), "Departments", DepartmentDto.class);
        List<ProjectDto> readProjects = NinjaExcel.readSheet(testFile.toFile(), "Projects", ProjectDto.class);

        assertThat(readUsers).hasSize(3);
        assertThat(readUsers.get(0).getName()).isEqualTo("Alice Johnson");
        assertThat(readUsers.get(1).getAge()).isEqualTo(25);
        assertThat(readUsers.get(2).getEmail()).isEqualTo("charlie@company.com");

        assertThat(readDepartments).hasSize(3);
        assertThat(readDepartments.get(0).getName()).isEqualTo("Engineering");
        assertThat(readDepartments.get(1).getBudget().compareTo(new BigDecimal("300000"))).isEqualTo(0);
        assertThat(readDepartments.get(2).getHeadCount()).isEqualTo(12);

        assertThat(readProjects).hasSize(2);
        assertThat(readProjects.get(0).getName()).isEqualTo("Project Alpha");
        assertThat(readProjects.get(0).getStartDate()).isEqualTo(LocalDate.of(2024, 1, 15));
        assertThat(readProjects.get(1).getBudget().compareTo(new BigDecimal("200000"))).isEqualTo(0);
    }

    @Test
    @DisplayName("ByteArrayOutputStream을 통한 다중 시트 생성 및 POI 검증")
    void createMultipleSheetsWithByteArrayOutput() throws Exception {
        List<UserDto> engineeringTeam = Arrays.asList(
                new UserDto(1L, "Alice", 30, "alice@tech.com"),
                new UserDto(2L, "Bob", 28, "bob@tech.com")
        );

        List<UserDto> marketingTeam = Arrays.asList(
                new UserDto(3L, "Carol", 32, "carol@marketing.com"),
                new UserDto(4L, "David", 29, "david@marketing.com")
        );

        ExcelWorkbook workbook = ExcelWorkbook.builder()
                .sheet("Engineering", engineeringTeam)
                .sheet("Marketing", marketingTeam)
                .build();

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        NinjaExcel.write(workbook, outputStream);
        byte[] excelBytes = outputStream.toByteArray();

        assertThat(excelBytes.length).isGreaterThan(0);

        try (Workbook poiWorkbook = WorkbookFactory.create(new ByteArrayInputStream(excelBytes))) {
            assertThat(poiWorkbook.getNumberOfSheets()).isEqualTo(2);
            assertThat(poiWorkbook.getSheetName(0)).isEqualTo("Engineering");
            assertThat(poiWorkbook.getSheetName(1)).isEqualTo("Marketing");

            Sheet engineeringSheet = poiWorkbook.getSheet("Engineering");
            assertThat(engineeringSheet.getPhysicalNumberOfRows()).isEqualTo(3);

            Row headerRow = engineeringSheet.getRow(0);
            assertThat(headerRow.getCell(0).getStringCellValue()).isEqualTo("ID");
            assertThat(headerRow.getCell(1).getStringCellValue()).isEqualTo("Name");
            assertThat(headerRow.getCell(2).getStringCellValue()).isEqualTo("Age");
            assertThat(headerRow.getCell(3).getStringCellValue()).isEqualTo("Email");

            Row firstDataRow = engineeringSheet.getRow(1);
            assertThat(firstDataRow.getCell(0).getNumericCellValue()).isEqualTo(1.0);
            assertThat(firstDataRow.getCell(1).getStringCellValue()).isEqualTo("Alice");
            assertThat(firstDataRow.getCell(2).getNumericCellValue()).isEqualTo(30.0);
            assertThat(firstDataRow.getCell(3).getStringCellValue()).isEqualTo("alice@tech.com");

            Sheet marketingSheet = poiWorkbook.getSheet("Marketing");
            assertThat(marketingSheet.getPhysicalNumberOfRows()).isEqualTo(3);

            Row marketingFirstRow = marketingSheet.getRow(1);
            assertThat(marketingFirstRow.getCell(1).getStringCellValue()).isEqualTo("Carol");
            assertThat(marketingFirstRow.getCell(3).getStringCellValue()).isEqualTo("carol@marketing.com");
        }
    }

    @Test
    @DisplayName("같은 타입의 여러 시트를 모두 읽기 테스트")
    void readAllSheetsWithSameType() throws Exception {
        List<UserDto> q1Users = Arrays.asList(
                new UserDto(1L, "Q1 User1", 25, "q1user1@company.com"),
                new UserDto(2L, "Q1 User2", 30, "q1user2@company.com")
        );

        List<UserDto> q2Users = Arrays.asList(
                new UserDto(3L, "Q2 User1", 28, "q2user1@company.com"),
                new UserDto(4L, "Q2 User2", 33, "q2user2@company.com"),
                new UserDto(5L, "Q2 User3", 27, "q2user3@company.com")
        );

        List<UserDto> q3Users = Collections.singletonList(new UserDto(6L, "Q3 User1", 31, "q3user1@company.com"));

        ExcelWorkbook workbook = ExcelWorkbook.builder()
                .sheet("Q1_Users", q1Users)
                .sheet("Q2_Users", q2Users)
                .sheet("Q3_Users", q3Users)
                .build();

        Path testFile = tempDir.resolve("quarterly_users.xlsx");
        NinjaExcel.write(workbook, testFile.toFile());

        Map<String, List<UserDto>> allQuarterlyUsers = NinjaExcel.readAllSheets(testFile.toFile(), UserDto.class);

        assertThat(allQuarterlyUsers).hasSize(3);
        assertThat(allQuarterlyUsers.get("Q1_Users")).hasSize(2);
        assertThat(allQuarterlyUsers.get("Q2_Users")).hasSize(3);
        assertThat(allQuarterlyUsers.get("Q3_Users")).hasSize(1);

        assertThat(allQuarterlyUsers.get("Q1_Users").get(0).getName()).isEqualTo("Q1 User1");
        assertThat(allQuarterlyUsers.get("Q2_Users").get(1).getName()).isEqualTo("Q2 User2");
        assertThat(allQuarterlyUsers.get("Q3_Users").get(0).getAge()).isEqualTo(31);
    }

    @Test
    @DisplayName("특정 시트들만 선택해서 읽기 테스트")
    void readSpecificSheets() throws Exception {
        List<UserDto> jan = Collections.singletonList(new UserDto(1L, "Jan User", 25, "jan@company.com"));
        List<UserDto> feb = Collections.singletonList(new UserDto(2L, "Feb User", 30, "feb@company.com"));
        List<UserDto> mar = Collections.singletonList(new UserDto(3L, "Mar User", 28, "mar@company.com"));
        List<UserDto> apr = Collections.singletonList(new UserDto(4L, "Apr User", 33, "apr@company.com"));

        ExcelWorkbook workbook = ExcelWorkbook.builder()
                .sheet("January", jan)
                .sheet("February", feb)
                .sheet("March", mar)
                .sheet("April", apr)
                .build();

        Path testFile = tempDir.resolve("monthly_users.xlsx");
        NinjaExcel.write(workbook, testFile.toFile());

        List<String> q1Months = Arrays.asList("January", "February", "March");
        Map<String, List<UserDto>> q1Data = NinjaExcel.readSheets(testFile.toFile(), UserDto.class, q1Months);

        assertThat(q1Data).hasSize(3);
        assertThat(q1Data).containsOnlyKeys("January", "February", "March");
        assertThat(q1Data).doesNotContainKey("April");

        assertThat(q1Data.get("January").get(0).getName()).isEqualTo("Jan User");
        assertThat(q1Data.get("February").get(0).getName()).isEqualTo("Feb User");
        assertThat(q1Data.get("March").get(0).getName()).isEqualTo("Mar User");
    }

    @Test
    @DisplayName("대상 시트만 읽는 API는 깨진 비대상 시트를 파싱하지 않는다")
    void targetedSheetApisSkipInvalidUnrequestedSheets() throws Exception {
        Path testFile = tempDir.resolve("partial_sheet_access.xlsx");
        long originalThreshold = NinjaExcel.getStreamingThreshold();

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet usersSheet = workbook.createSheet("Users");
            Row usersHeader = usersSheet.createRow(0);
            usersHeader.createCell(0).setCellValue("ID");
            usersHeader.createCell(1).setCellValue("Name");
            usersHeader.createCell(2).setCellValue("Age");
            usersHeader.createCell(3).setCellValue("Email");

            Row usersRow = usersSheet.createRow(1);
            usersRow.createCell(0).setCellValue(1);
            usersRow.createCell(1).setCellValue("Alice");
            usersRow.createCell(2).setCellValue(30);
            usersRow.createCell(3).setCellValue("alice@example.com");

            Sheet brokenSheet = workbook.createSheet("Broken");
            Row brokenHeader = brokenSheet.createRow(0);
            brokenHeader.createCell(0).setCellValue("ID");
            brokenHeader.createCell(2).setCellValue("Name");

            try (OutputStream outputStream = java.nio.file.Files.newOutputStream(testFile)) {
                workbook.write(outputStream);
            }
        }

        try {
            for (long threshold : new long[]{Long.MAX_VALUE, 1L}) {
                NinjaExcel.setStreamingThreshold(threshold);

                List<UserDto> firstSheetRows = NinjaExcel.read(testFile.toFile(), UserDto.class);
                List<UserDto> namedRows = NinjaExcel.readSheet(testFile.toFile(), "Users", UserDto.class);
                Map<String, List<UserDto>> selectedRows = NinjaExcel.readSheets(
                        testFile.toFile(),
                        UserDto.class,
                        Collections.singletonList("Users")
                );
                List<String> sheetNames = NinjaExcel.getSheetNames(testFile.toFile());

                assertThat(firstSheetRows).hasSize(1);
                assertThat(firstSheetRows.get(0).getName()).isEqualTo("Alice");
                assertThat(namedRows).hasSize(1);
                assertThat(namedRows.get(0).getEmail()).isEqualTo("alice@example.com");
                assertThat(selectedRows).containsOnlyKeys("Users");
                assertThat(selectedRows.get("Users")).hasSize(1);
                assertThat(sheetNames).containsExactly("Users", "Broken");
            }
        } finally {
            NinjaExcel.setStreamingThreshold(originalThreshold);
        }
    }

    @Test
    @DisplayName("워크북 메타데이터와 함께 다중 시트 처리")
    void multiSheetsWithMetadata() throws Exception {
        List<UserDto> employees = Arrays.asList(
                new UserDto(1L, "John Doe", 35, "john@company.com")
        );

        List<DepartmentDto> departments = Arrays.asList(
                new DepartmentDto(1L, "IT", new BigDecimal("1000000"), 20)
        );

        WorkbookMetadata metadata = new WorkbookMetadata(
                "HR Department",
                "Annual Company Report 2024",
                LocalDateTime.of(2024, 12, 31, 23, 59, 59)
        );

        ExcelWorkbook workbook = ExcelWorkbook.builder()
                .sheet("Employees", employees)
                .sheet("Departments", departments)
                .metadata(metadata)
                .build();

        Path testFile = tempDir.resolve("company_report_with_metadata.xlsx");
        NinjaExcel.write(workbook, testFile.toFile());

        assertThat(testFile.toFile()).exists();
        assertThat(testFile.toFile().length()).isGreaterThan(0);

        try (Workbook poiWorkbook = WorkbookFactory.create(testFile.toFile())) {
            org.apache.poi.xssf.usermodel.XSSFWorkbook xssfWorkbook = (org.apache.poi.xssf.usermodel.XSSFWorkbook) poiWorkbook;
            assertThat(xssfWorkbook.getProperties().getCoreProperties().getCreator()).isEqualTo(metadata.getAuthor());
            assertThat(xssfWorkbook.getProperties().getCoreProperties().getTitle()).isEqualTo(metadata.getTitle());
        }

        List<UserDto> readEmployees = NinjaExcel.readSheet(testFile.toFile(), "Employees", UserDto.class);
        List<DepartmentDto> readDepartments = NinjaExcel.readSheet(testFile.toFile(), "Departments", DepartmentDto.class);
        ExcelWorkbook poiReadWorkbook = new PoiWorkbookReader().read(testFile.toFile());
        ExcelWorkbook streamingReadWorkbook = new StreamingWorkbookReader().read(testFile.toFile());

        assertThat(readEmployees).hasSize(1);
        assertThat(readDepartments).hasSize(1);
        assertThat(readEmployees.get(0).getName()).isEqualTo("John Doe");
        assertThat(readDepartments.get(0).getName()).isEqualTo("IT");
        assertThat(poiReadWorkbook.getMetadata().getAuthor()).isEqualTo(metadata.getAuthor());
        assertThat(poiReadWorkbook.getMetadata().getTitle()).isEqualTo(metadata.getTitle());
        assertThat(poiReadWorkbook.getMetadata().getCreatedDate()).isEqualTo(metadata.getCreatedDate());
        assertThat(streamingReadWorkbook.getMetadata().getAuthor()).isEqualTo(metadata.getAuthor());
        assertThat(streamingReadWorkbook.getMetadata().getTitle()).isEqualTo(metadata.getTitle());
        assertThat(streamingReadWorkbook.getMetadata().getCreatedDate()).isEqualTo(metadata.getCreatedDate());
    }

    @Test
    @DisplayName("메타데이터를 지정하지 않으면 createdDate를 임의 생성하지 않는다")
    void workbookWithoutMetadataDoesNotFabricateCreatedDate() throws Exception {
        List<UserDto> employees = Arrays.asList(
                new UserDto(1L, "John Doe", 35, "john@company.com")
        );

        ExcelWorkbook workbook = ExcelWorkbook.builder()
                .sheet("Employees", employees)
                .build();

        assertThat(workbook.getMetadata().getAuthor()).isNull();
        assertThat(workbook.getMetadata().getTitle()).isNull();
        assertThat(workbook.getMetadata().getCreatedDate()).isNull();

        Path testFile = tempDir.resolve("company_report_without_metadata.xlsx");
        NinjaExcel.write(workbook, testFile.toFile());

        ExcelWorkbook poiReadWorkbook = new PoiWorkbookReader().read(testFile.toFile());
        ExcelWorkbook streamingReadWorkbook = new StreamingWorkbookReader().read(testFile.toFile());

        assertThat(poiReadWorkbook.getMetadata().getAuthor()).isNull();
        assertThat(poiReadWorkbook.getMetadata().getTitle()).isNull();
        assertThat(poiReadWorkbook.getMetadata().getCreatedDate()).isNull();
        assertThat(streamingReadWorkbook.getMetadata().getAuthor()).isNull();
        assertThat(streamingReadWorkbook.getMetadata().getTitle()).isNull();
        assertThat(streamingReadWorkbook.getMetadata().getCreatedDate()).isNull();
    }

    @Test
    @DisplayName("빈 시트와 데이터가 있는 시트 혼합 처리")
    void mixEmptyAndNonEmptySheets() throws Exception {
        List<UserDto> activeUsers = Arrays.asList(
                new UserDto(1L, "Active User", 30, "active@company.com")
        );

        List<UserDto> emptyUsers = Collections.emptyList();

        List<UserDto> inactiveUsers = Arrays.asList(
                new UserDto(999L, "No Data", 0, "nodata@company.com")
        );

        ExcelWorkbook workbook = ExcelWorkbook.builder()
                .sheet("ActiveUsers", activeUsers)
                .sheet("InactiveUsers", inactiveUsers)
                .build();

        Path testFile = tempDir.resolve("mixed_sheets.xlsx");
        NinjaExcel.write(workbook, testFile.toFile());

        List<String> sheetNames = NinjaExcel.getSheetNames(testFile.toFile());
        assertThat(sheetNames).hasSize(2);

        List<UserDto> activeData = NinjaExcel.readSheet(testFile.toFile(), "ActiveUsers", UserDto.class);
        List<UserDto> inactiveData = NinjaExcel.readSheet(testFile.toFile(), "InactiveUsers", UserDto.class);

        assertThat(activeData).hasSize(1);
        assertThat(inactiveData).hasSize(1);
        assertThat(activeData.get(0).getName()).isEqualTo("Active User");
        assertThat(inactiveData.get(0).getId()).isEqualTo(999L);
    }
}
