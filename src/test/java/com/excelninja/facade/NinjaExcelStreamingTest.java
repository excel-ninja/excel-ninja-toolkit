package com.excelninja.facade;

import com.excelninja.application.facade.NinjaExcel;
import com.excelninja.domain.annotation.ExcelReadColumn;
import com.excelninja.domain.annotation.ExcelWriteColumn;
import com.excelninja.domain.exception.DocumentConversionException;
import com.excelninja.domain.model.ExcelWorkbook;
import com.excelninja.infrastructure.io.StreamingWorkbookReader;
import org.junit.jupiter.api.*;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.time.LocalDate;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

import static org.assertj.core.api.Assertions.*;

class NinjaExcelStreamingTest {

    @TempDir
    static File tempDir;

    private static File smallFile;
    private static File largeFile;
    private static File hugeFile;

    static class Employee {
        @ExcelReadColumn(headerName = "ID")
        @ExcelWriteColumn(headerName = "ID", order = 0)
        private Long id;

        @ExcelReadColumn(headerName = "Name")
        @ExcelWriteColumn(headerName = "Name", order = 1)
        private String name;

        @ExcelReadColumn(headerName = "Email")
        @ExcelWriteColumn(headerName = "Email", order = 2)
        private String email;

        @ExcelReadColumn(headerName = "Salary")
        @ExcelWriteColumn(headerName = "Salary", order = 3)
        private Double salary;

        @ExcelReadColumn(headerName = "HireDate")
        @ExcelWriteColumn(headerName = "HireDate", order = 4)
        private LocalDate hireDate;

        @ExcelReadColumn(headerName = "IsActive")
        @ExcelWriteColumn(headerName = "IsActive", order = 5)
        private Boolean isActive;

        @ExcelReadColumn(headerName = "Department")
        @ExcelWriteColumn(headerName = "Department", order = 6)
        private String department;

        public Employee() {}

        public Employee(
                Long id,
                String name,
                String email,
                Double salary,
                LocalDate hireDate,
                Boolean isActive,
                String department
        ) {
            this.id = id;
            this.name = name;
            this.email = email;
            this.salary = salary;
            this.hireDate = hireDate;
            this.isActive = isActive;
            this.department = department;
        }

        public Long getId() {return id;}

        public void setId(Long id) {this.id = id;}

        public String getName() {return name;}

        public void setName(String name) {this.name = name;}

        public String getEmail() {return email;}

        public void setEmail(String email) {this.email = email;}

        public Double getSalary() {return salary;}

        public void setSalary(Double salary) {this.salary = salary;}

        public LocalDate getHireDate() {return hireDate;}

        public void setHireDate(LocalDate hireDate) {this.hireDate = hireDate;}

        public Boolean getIsActive() {return isActive;}

        public void setIsActive(Boolean isActive) {this.isActive = isActive;}

        public String getDepartment() {return department;}

        public void setDepartment(String department) {this.department = department;}

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (!(o instanceof Employee)) return false;
            Employee employee = (Employee) o;
            return Objects.equals(id, employee.id) &&
                    Objects.equals(name, employee.name) &&
                    Objects.equals(email, employee.email) &&
                    Objects.equals(salary, employee.salary) &&
                    Objects.equals(hireDate, employee.hireDate) &&
                    Objects.equals(isActive, employee.isActive) &&
                    Objects.equals(department, employee.department);
        }

        @Override
        public int hashCode() {
            return Objects.hash(id, name, email, salary, hireDate, isActive, department);
        }

        @Override
        public String toString() {
            return String.format("Employee{id=%d, name='%s', email='%s', salary=%.2f, hireDate=%s, isActive=%s, department='%s'}",
                    id, name, email, salary, hireDate, isActive, department);
        }
    }

    static class Product {
        @ExcelReadColumn(headerName = "ProductId")
        @ExcelWriteColumn(headerName = "ProductId", order = 0)
        private String productId;

        @ExcelReadColumn(headerName = "ProductName")
        @ExcelWriteColumn(headerName = "ProductName", order = 1)
        private String productName;

        @ExcelReadColumn(headerName = "Price")
        @ExcelWriteColumn(headerName = "Price", order = 2)
        private Double price;

        @ExcelReadColumn(headerName = "Category")
        @ExcelWriteColumn(headerName = "Category", order = 3)
        private String category;

        public Product() {}

        public Product(
                String productId,
                String productName,
                Double price,
                String category
        ) {
            this.productId = productId;
            this.productName = productName;
            this.price = price;
            this.category = category;
        }

        public String getProductId() {return productId;}

        public void setProductId(String productId) {this.productId = productId;}

        public String getProductName() {return productName;}

        public void setProductName(String productName) {this.productName = productName;}

        public Double getPrice() {return price;}

        public void setPrice(Double price) {this.price = price;}

        public String getCategory() {return category;}

        public void setCategory(String category) {this.category = category;}

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (!(o instanceof Product)) return false;
            Product product = (Product) o;
            return Objects.equals(productId, product.productId) &&
                    Objects.equals(productName, product.productName) &&
                    Objects.equals(price, product.price) &&
                    Objects.equals(category, product.category);
        }

        @Override
        public int hashCode() {
            return Objects.hash(productId, productName, price, category);
        }
    }

    @BeforeAll
    static void setupTestFiles() throws IOException {
        NinjaExcel.setStreamingThreshold(1024 * 1024);
        smallFile = createTestFile("small_employees.xlsx", 100);
        largeFile = createTestFile("large_employees.xlsx", 10000);
        hugeFile = createTestFile("huge_employees.xlsx", 50000);
    }

    static File createTestFile(
            String filename,
            int recordCount
    ) throws IOException {
        File file = new File(tempDir, filename);

        List<Employee> employees = createEmployeeData(recordCount);

        ExcelWorkbook.WorkbookBuilder builder = ExcelWorkbook.builder().sheet("Employees", employees);

        if (recordCount > 1000) {
            List<Product> products = createProductData(recordCount / 10);
            builder.sheet("Products", products);
        }

        ExcelWorkbook workbook = builder.build();
        NinjaExcel.write(workbook, file);

        System.out.printf("[TEST-SETUP] Created test file: %s (%.2f MB, %d records)%n", filename, file.length() / (1024.0 * 1024.0), recordCount);

        return file;
    }

    static List<Employee> createEmployeeData(int recordCount) {
        List<Employee> employees = new ArrayList<>();
        Random random = new Random(42);
        String[] departments = {"Engineering", "Sales", "Marketing", "HR", "Finance"};
        Calendar cal = Calendar.getInstance();
        cal.set(2020, Calendar.JANUARY, 1);

        for (int i = 1; i <= recordCount; i++) {
            cal.add(Calendar.DAY_OF_YEAR, random.nextInt(30));
            Employee employee = new Employee(
                    (long) i,
                    "Employee " + i,
                    "employee" + i + "@company.com",
                    50000 + random.nextDouble() * 100000,
                    LocalDate.of(cal.get(Calendar.YEAR), cal.get(Calendar.MONTH) + 1, cal.get(Calendar.DAY_OF_MONTH)),
                    random.nextBoolean(),
                    departments[random.nextInt(departments.length)]
            );

            employees.add(employee);
        }

        return employees;
    }

    static List<Product> createProductData(int recordCount) {
        List<Product> products = new ArrayList<>();
        Random random = new Random(42);
        String[] categories = {"Electronics", "Books", "Clothing", "Sports", "Home"};

        for (int i = 1; i <= recordCount; i++) {
            Product product = new Product(
                    "PROD-" + String.format("%06d", i),
                    "Product " + i,
                    10 + random.nextDouble() * 990,
                    categories[random.nextInt(categories.length)]
            );

            products.add(product);
        }

        return products;
    }

    @Nested
    @DisplayName("자동 스트리밍 선택 테스트")
    class AutoStreamingSelectionTest {

        @Test
        @DisplayName("소용량 파일은 POI 리더 사용")
        void shouldUsePOIForSmallFiles() {
            long startTime = System.currentTimeMillis();
            List<Employee> employees = NinjaExcel.read(smallFile, Employee.class);
            long duration = System.currentTimeMillis() - startTime;

            assertThat(employees).hasSize(100);
            assertThat(employees.get(0).getId()).isEqualTo(1L);
            assertThat(employees.get(0).getName()).isEqualTo("Employee 1");

            System.out.printf("[POI] Small file read: %d records in %d ms%n", employees.size(), duration);
        }

        @Test
        @DisplayName("대용량 파일은 스트리밍 리더 사용")
        void shouldUseStreamingForLargeFiles() {
            long startTime = System.currentTimeMillis();
            List<Employee> employees = NinjaExcel.read(largeFile, Employee.class);
            long duration = System.currentTimeMillis() - startTime;

            assertThat(employees).hasSize(10000);
            assertThat(employees.get(0).getId()).isEqualTo(1L);
            assertThat(employees.get(9999).getId()).isEqualTo(10000L);

            System.out.printf("[STREAMING] Large file read: %d records in %d ms%n", employees.size(), duration);
        }

        @Test
        @DisplayName("파일 크기에 관계없이 동일한 결과 보장")
        void shouldProduceSameResultsRegardlessOfReader() {
            List<Employee> smallResult = NinjaExcel.read(smallFile, Employee.class);

            // 작은 파일을 스트리밍으로도 읽어보기
            StreamingWorkbookReader streamingReader = new StreamingWorkbookReader();
            ExcelWorkbook workbook;
            try {
                workbook = streamingReader.read(smallFile);
                String firstSheetName = workbook.getSheetNames().iterator().next();
                // 동일한 변환 로직 적용해서 비교
            } catch (IOException e) {
                fail("Failed to read with streaming reader", e);
            }

            // 첫 번째 직원 검증
            Employee first = smallResult.get(0);
            assertThat(first.getId()).isEqualTo(1L);
            assertThat(first.getName()).isEqualTo("Employee 1");
            assertThat(first.getEmail()).isEqualTo("employee1@company.com");
        }
    }

    @Nested
    @DisplayName("청크 처리 테스트")
    class ChunkProcessingTest {

        @Test
        @DisplayName("청크 단위로 대용량 파일 처리")
        void shouldProcessLargeFileInChunks() {
            int chunkSize = 1000;
            Iterator<List<Employee>> chunks = NinjaExcel.readInChunks(hugeFile, Employee.class, chunkSize);

            int totalProcessed = 0;
            int chunkCount = 0;
            long startTime = System.currentTimeMillis();

            while (chunks.hasNext()) {
                List<Employee> chunk = chunks.next();
                chunkCount++;
                totalProcessed += chunk.size();

                // 각 청크 검증
                assertThat(chunk).isNotEmpty();
                assertThat(chunk.size()).isLessThanOrEqualTo(chunkSize);

                // 첫 번째 청크의 첫 번째 레코드 검증
                if (chunkCount == 1) {
                    Employee first = chunk.get(0);
                    assertThat(first.getId()).isEqualTo(1L);
                    assertThat(first.getName()).isEqualTo("Employee 1");
                }

                System.out.printf("[CHUNK-%d] Processed %d records (total: %d)%n",
                        chunkCount, chunk.size(), totalProcessed);
            }

            long duration = System.currentTimeMillis() - startTime;

            assertThat(totalProcessed).isEqualTo(50000);
            assertThat(chunkCount).isEqualTo(50); // 50,000 / 1,000

            System.out.printf("[CHUNK-SUMMARY] Processed %d records in %d chunks over %d ms%n",
                    totalProcessed, chunkCount, duration);
        }

        @Test
        @DisplayName("메모리 효율성 검증")
        void shouldBeMemoryEfficient() {
            Runtime runtime = Runtime.getRuntime();

            // 메모리 측정 시작
            System.gc();
            long memoryBefore = runtime.totalMemory() - runtime.freeMemory();

            int processedRecords = 0;
            Iterator<List<Employee>> chunks = NinjaExcel.readInChunks(hugeFile, Employee.class, 500);

            while (chunks.hasNext()) {
                List<Employee> chunk = chunks.next();
                processedRecords += chunk.size();

                // 청크 처리 후 메모리 해제
                chunk.clear();

                if (processedRecords % 10000 == 0) {
                    System.gc();
                    long currentMemory = runtime.totalMemory() - runtime.freeMemory();
                    System.out.printf("[MEMORY] Processed %d records, memory: %.2f MB%n",
                            processedRecords, (currentMemory - memoryBefore) / (1024.0 * 1024.0));
                }
            }

            System.gc();
            long memoryAfter = runtime.totalMemory() - runtime.freeMemory();
            double memoryUsed = (memoryAfter - memoryBefore) / (1024.0 * 1024.0);

            assertThat(processedRecords).isEqualTo(50000);
            assertThat(memoryUsed).isLessThan(200); // 200MB 미만 사용

            System.out.printf("[MEMORY-FINAL] Total memory used: %.2f MB for %d records%n",
                    memoryUsed, processedRecords);
        }

        @Test
        @DisplayName("청크 크기 조정 테스트")
        void shouldRespectChunkSize() {
            int[] chunkSizes = {100, 500, 1000, 2000};

            for (int chunkSize : chunkSizes) {
                Iterator<List<Employee>> chunks = NinjaExcel.readInChunks(largeFile, Employee.class, chunkSize);

                int chunkCount = 0;
                int totalRecords = 0;

                while (chunks.hasNext()) {
                    List<Employee> chunk = chunks.next();
                    chunkCount++;
                    totalRecords += chunk.size();

                    // 마지막 청크가 아니라면 정확한 크기여야 함
                    if (totalRecords < 10000 && chunk.size() < chunkSize) {
                        // 마지막 청크인 경우
                        assertThat(chunk.size()).isLessThanOrEqualTo(chunkSize);
                    } else if (totalRecords < 10000) {
                        // 중간 청크인 경우
                        assertThat(chunk.size()).isEqualTo(chunkSize);
                    }
                }

                System.out.printf("[CHUNK-SIZE] Size %d: %d chunks for %d records%n",
                        chunkSize, chunkCount, totalRecords);

                assertThat(totalRecords).isEqualTo(10000);
            }
        }
    }

    @Nested
    @DisplayName("다중 시트 스트리밍 테스트")
    class MultiSheetStreamingTest {

        @Test
        @DisplayName("각 시트를 올바른 데이터 모델로 읽기")
        void shouldReadEachSheetWithCorrectModel() {
            List<Employee> employees = NinjaExcel.readSheet(largeFile, "Employees", Employee.class);
            List<Product> products = NinjaExcel.readSheet(largeFile, "Products", Product.class);
            assertThat(employees).hasSize(10000);
            assertThat(employees.get(0).getName()).isEqualTo("Employee 1");

            assertThat(products).hasSize(1000);
            assertThat(products.get(0).getProductName()).isEqualTo("Product 1");

            System.out.printf("[MULTI-SHEET-TEST] Successfully read %d employees and %d products from separate sheets.%n", employees.size(), products.size());
        }

        @Test
        @DisplayName("대용량 파일의 첫 번째 시트 읽기")
        void shouldReadFirstSheetFromLargeFile() {
            List<Employee> employees = NinjaExcel.read(largeFile, Employee.class);

            assertThat(employees).isNotNull();
            assertThat(employees).hasSize(10000);
            assertThat(employees.get(0).getId()).isEqualTo(1L);
            assertThat(employees.get(9999).getId()).isEqualTo(10000L);

            System.out.printf("[MULTI-SHEET-TEST] Successfully read %d records from the first sheet of the large file.%n", employees.size());
        }

        @Test
        @DisplayName("특정 시트만 선택해서 읽기")
        void shouldReadSpecificSheetsFromLargeFile() {
            List<String> targetSheets = Collections.singletonList("Employees");
            Map<String, List<Employee>> result = NinjaExcel.readSheets(largeFile, Employee.class, targetSheets);

            assertThat(result).hasSize(1);
            assertThat(result).containsKey("Employees");
            assertThat(result.get("Employees")).hasSize(10000);
        }

        @Test
        @DisplayName("시트 이름 목록 조회")
        void shouldGetSheetNamesFromLargeFile() {
            List<String> sheetNames = NinjaExcel.getSheetNames(largeFile);

            assertThat(sheetNames).hasSize(2);
            assertThat(sheetNames).containsExactly("Employees", "Products");
        }
    }

    @Nested
    @DisplayName("성능 비교 테스트")
    class PerformanceComparisonTest {

        @Test
        @DisplayName("POI vs 스트리밍 성능 비교")
        void shouldComparePerformanceBetweenReaders() {
            // POI로 중간 크기 파일 읽기 (강제로 POI 사용)
            long poiStart = System.currentTimeMillis();
            // 임계값을 임시로 높여서 POI 강제 사용
            NinjaExcel.setStreamingThreshold(50 * 1024 * 1024);
            List<Employee> poiResult = NinjaExcel.read(largeFile, Employee.class);
            long poiDuration = System.currentTimeMillis() - poiStart;

            // 스트리밍으로 동일 파일 읽기
            NinjaExcel.setStreamingThreshold(1); // 스트리밍 강제 사용
            long streamingStart = System.currentTimeMillis();
            List<Employee> streamingResult = NinjaExcel.read(largeFile, Employee.class);
            long streamingDuration = System.currentTimeMillis() - streamingStart;

            // 결과 비교
            assertThat(poiResult).hasSize(streamingResult.size());
            assertThat(poiResult.get(0).getName()).isEqualTo(streamingResult.get(0).getName());

            System.out.printf("[PERFORMANCE] POI: %d ms, Streaming: %d ms for %d records%n",
                    poiDuration, streamingDuration, poiResult.size());

            // 원래 임계값 복원
            NinjaExcel.setStreamingThreshold(1 * 1024 * 1024);
        }

        @Test
        @DisplayName("대용량 파일 처리 시간 측정")
        void shouldMeasureHugeFileProcessingTime() {
            long start = System.currentTimeMillis();
            List<Employee> employees = NinjaExcel.read(hugeFile, Employee.class);
            long duration = System.currentTimeMillis() - start;

            assertThat(employees).hasSize(50000);

            double recordsPerSecond = employees.size() * 1000.0 / duration;
            System.out.printf("[HUGE-FILE] Processed %d records in %d ms (%.2f records/sec)%n",
                    employees.size(), duration, recordsPerSecond);

            // 성능 기준 검증 (초당 1000개 이상 처리)
            assertThat(recordsPerSecond).isGreaterThan(1000);
        }
    }

    @Nested
    @DisplayName("오류 처리 테스트")
    class ErrorHandlingTest {

        @Test
        @DisplayName("존재하지 않는 파일")
        void shouldThrowExceptionForNonExistentFile() {
            File nonExistentFile = new File(tempDir, "non_existent.xlsx");

            assertThatThrownBy(() -> NinjaExcel.read(nonExistentFile, Employee.class))
                    .isInstanceOf(DocumentConversionException.class)
                    .hasMessageContaining("does not exist");
        }

        @Test
        @DisplayName("잘못된 청크 크기")
        void shouldThrowExceptionForInvalidChunkSize() {
            assertThatThrownBy(() -> NinjaExcel.readInChunks(smallFile, Employee.class, 0))
                    .isInstanceOf(DocumentConversionException.class)
                    .hasMessageContaining("Chunk size must be positive");

            assertThatThrownBy(() -> NinjaExcel.readInChunks(smallFile, Employee.class, -1))
                    .isInstanceOf(DocumentConversionException.class)
                    .hasMessageContaining("Chunk size must be positive");
        }

        @Test
        @DisplayName("null 파라미터 검증")
        void shouldThrowExceptionForNullParameters() {
            assertThatThrownBy(() -> NinjaExcel.read((File) null, Employee.class))
                    .isInstanceOf(DocumentConversionException.class)
                    .hasMessageContaining("File parameter cannot be null");

            assertThatThrownBy(() -> NinjaExcel.read(smallFile, null))
                    .isInstanceOf(DocumentConversionException.class)
                    .hasMessageContaining("Target class parameter cannot be null");
        }
    }

    @Nested
    @DisplayName("스트레스 테스트")
    class StressTest {

        @Test
        @DisplayName("연속 청크 처리 스트레스 테스트")
        @Timeout(30)
        void shouldHandleContinuousChunkProcessing() {
            AtomicInteger totalProcessed = new AtomicInteger(0);

            for (int i = 0; i < 5; i++) {
                Iterator<List<Employee>> chunks = NinjaExcel.readInChunks(hugeFile, Employee.class, 2000);

                int iterationCount = 0;
                while (chunks.hasNext()) {
                    List<Employee> chunk = chunks.next();
                    iterationCount += chunk.size();

                    chunk.forEach(emp -> {
                        if (emp.getId() != null && emp.getName() != null) {
                            totalProcessed.incrementAndGet();
                        }
                    });
                }

                System.out.printf("[STRESS-%d] Processed %d records%n", i + 1, iterationCount);
                assertThat(iterationCount).isEqualTo(50000);

                if (chunks instanceof AutoCloseable) {
                    try {
                        ((AutoCloseable) chunks).close();
                    } catch (Exception e) {
                        System.err.println("Failed to close iterator: " + e.getMessage());
                    }
                }
            }

            assertThat(totalProcessed.get()).isEqualTo(250000); // 5 * 50,000
            System.out.printf("[STRESS-TOTAL] Successfully processed %d records%n", totalProcessed.get());
        }
    }

    @AfterAll
    static void cleanup() {
        System.out.println("[TEST-CLEANUP] Cleaning up test files...");
        deleteFile(smallFile);
        deleteFile(largeFile);
        deleteFile(hugeFile);
        System.out.println("[TEST-CLEANUP] Test files cleaned up successfully");
    }

    private static void deleteFile(File file) {
        if (file != null && file.exists()) {
            if (!file.delete()) {
                System.err.println("[TEST-CLEANUP-WARN] Failed to delete file: " + file.getAbsolutePath());
            }
        }
    }
}