package com.excelninja.facade;

import com.excelninja.application.facade.NinjaExcel;
import com.excelninja.domain.annotation.ExcelReadColumn;
import com.excelninja.domain.annotation.ExcelWriteColumn;
import com.excelninja.domain.model.ExcelWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;

import static org.assertj.core.api.Assertions.assertThat;

@DisplayName("NinjaExcel concurrency")
class NinjaExcelConcurrencyTest {

    @TempDir
    Path tempDir;

    static class EmployeeRow {
        @ExcelReadColumn(headerName = "ID")
        @ExcelWriteColumn(headerName = "ID", order = 0)
        private Long id;

        @ExcelReadColumn(headerName = "Name")
        @ExcelWriteColumn(headerName = "Name", order = 1)
        private String name;

        @ExcelReadColumn(headerName = "HireDate")
        @ExcelWriteColumn(headerName = "HireDate", order = 2)
        private LocalDate hireDate;

        public EmployeeRow() {}

        EmployeeRow(Long id, String name, LocalDate hireDate) {
            this.id = id;
            this.name = name;
            this.hireDate = hireDate;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (!(o instanceof EmployeeRow)) return false;
            EmployeeRow that = (EmployeeRow) o;
            return Objects.equals(id, that.id)
                    && Objects.equals(name, that.name)
                    && Objects.equals(hireDate, that.hireDate);
        }

        @Override
        public int hashCode() {
            return Objects.hash(id, name, hireDate);
        }
    }

    @Test
    @DisplayName("Concurrent POI reads return identical results")
    void shouldReadSameWorkbookConcurrentlyWithPoiReader() throws Exception {
        File workbookFile = createWorkbookFile("poi_concurrent.xlsx", 120);
        long originalThreshold = NinjaExcel.getStreamingThreshold();
        NinjaExcel.setStreamingThreshold(Long.MAX_VALUE);

        try {
            List<List<EmployeeRow>> results = runConcurrentReads(workbookFile, 8);

            assertThat(results).hasSize(8);
            assertThat(results).allSatisfy(rows -> assertThat(rows).hasSize(120));
            assertThat(results).allMatch(rows -> rows.equals(results.get(0)));
        } finally {
            NinjaExcel.setStreamingThreshold(originalThreshold);
        }
    }

    @Test
    @DisplayName("Concurrent streaming reads return identical results")
    void shouldReadSameWorkbookConcurrentlyWithStreamingReader() throws Exception {
        File workbookFile = createWorkbookFile("streaming_concurrent.xlsx", 120);
        long originalThreshold = NinjaExcel.getStreamingThreshold();
        NinjaExcel.setStreamingThreshold(1);

        try {
            List<List<EmployeeRow>> results = runConcurrentReads(workbookFile, 8);

            assertThat(results).hasSize(8);
            assertThat(results).allSatisfy(rows -> assertThat(rows).hasSize(120));
            assertThat(results).allMatch(rows -> rows.equals(results.get(0)));
        } finally {
            NinjaExcel.setStreamingThreshold(originalThreshold);
        }
    }

    @Test
    @DisplayName("Concurrent writes produce readable workbooks")
    void shouldWriteWorkbooksConcurrently() throws Exception {
        List<EmployeeRow> employees = createEmployees(40);
        ExcelWorkbook workbook = ExcelWorkbook.builder().sheet("Employees", employees).build();
        ExecutorService executorService = Executors.newFixedThreadPool(6);

        try {
            List<Callable<List<EmployeeRow>>> tasks = new ArrayList<>();
            for (int i = 0; i < 6; i++) {
                final int taskIndex = i;
                tasks.add(() -> {
                    File outputFile = tempDir.resolve("concurrent-write-" + taskIndex + ".xlsx").toFile();
                    NinjaExcel.write(workbook, outputFile);
                    return NinjaExcel.read(outputFile, EmployeeRow.class);
                });
            }

            List<Future<List<EmployeeRow>>> futures = executorService.invokeAll(tasks);
            for (Future<List<EmployeeRow>> future : futures) {
                assertThat(future.get()).isEqualTo(employees);
            }
        } finally {
            executorService.shutdownNow();
        }
    }

    private List<List<EmployeeRow>> runConcurrentReads(File workbookFile, int taskCount) throws Exception {
        ExecutorService executorService = Executors.newFixedThreadPool(taskCount);

        try {
            List<Callable<List<EmployeeRow>>> tasks = new ArrayList<>();
            for (int i = 0; i < taskCount; i++) {
                tasks.add(() -> NinjaExcel.read(workbookFile, EmployeeRow.class));
            }

            List<Future<List<EmployeeRow>>> futures = executorService.invokeAll(tasks);
            List<List<EmployeeRow>> results = new ArrayList<>();
            for (Future<List<EmployeeRow>> future : futures) {
                results.add(future.get());
            }
            return results;
        } finally {
            executorService.shutdownNow();
        }
    }

    private File createWorkbookFile(String fileName, int recordCount) {
        File workbookFile = tempDir.resolve(fileName).toFile();
        ExcelWorkbook workbook = ExcelWorkbook.builder()
                .sheet("Employees", createEmployees(recordCount))
                .build();
        NinjaExcel.write(workbook, workbookFile);
        return workbookFile;
    }

    private List<EmployeeRow> createEmployees(int recordCount) {
        List<EmployeeRow> employees = new ArrayList<>();
        for (int i = 1; i <= recordCount; i++) {
            employees.add(new EmployeeRow((long) i, "Employee " + i, LocalDate.of(2024, 1, 1).plusDays(i)));
        }
        return employees;
    }
}
