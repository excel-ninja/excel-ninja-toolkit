package com.excelninja.domain.model;

import com.excelninja.application.port.ConverterPort;
import com.excelninja.domain.annotation.ExcelWriteColumn;
import com.excelninja.domain.annotation.ExcelReadColumn;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelDocument {
    private final String sheetName;
    private final List<String> headers;
    private final List<List<Object>> rows;
    private final Map<Integer, Integer> columnWidths;
    private final Map<Integer, Short> rowHeights;

    private ExcelDocument(
            String sheetName,
            List<String> headers,
            List<List<Object>> rows,
            Map<Integer, Integer> columnWidths,
            Map<Integer, Short> rowHeights
    ) {
        this.sheetName = Objects.requireNonNull(sheetName);
        this.headers = List.copyOf(Objects.requireNonNull(headers));
        this.rows = List.copyOf(Objects.requireNonNull(rows));
        this.columnWidths = Map.copyOf(Objects.requireNonNull(columnWidths));
        this.rowHeights = Map.copyOf(Objects.requireNonNull(rowHeights));
    }

    public String getSheetName() {
        return this.sheetName;
    }

    public List<String> getHeaders() {
        return this.headers;
    }

    public List<List<Object>> getRows() {
        return this.rows;
    }

    public Map<Integer, Integer> getColumnWidths() {
        return this.columnWidths;
    }

    public Map<Integer, Short> getRowHeights() {
        return this.rowHeights;
    }

    public static ExcelDocument.Builder builder() {
        return new ExcelDocument.Builder();
    }

    public static final class Builder {
        private String sheetName;
        private final List<String> headers = new ArrayList<>();
        private final List<List<Object>> rows = new ArrayList<>();
        private final Map<Integer, Integer> columnWidths = new HashMap<>();
        private final Map<Integer, Short> rowHeights = new HashMap<>();

        public Builder() {}

        public Builder sheet(String sheetName) {
            this.sheetName = Objects.requireNonNull(sheetName);
            return this;
        }

        public Builder headers(List<String> headers) {
            Objects.requireNonNull(headers);
            this.headers.clear();
            this.headers.addAll(headers);
            return this;
        }

        public Builder rows(List<List<Object>> rowValues) {
            for (var rowValue : rowValues) {
                Objects.requireNonNull(rowValue);
                if (rowValue.size() != headers.size()) {
                    throw new IllegalArgumentException("Row values size (" + rowValue.size() + ") does not match header count (" + headers.size() + ")");
                }
                this.rows.add(new ArrayList<>(rowValue));
            }
            return this;
        }
        public ExcelDocument build() {
            return new ExcelDocument(sheetName, headers, rows, columnWidths, rowHeights);
        }
    }

    public static <T> ExcelDocument createWriter(
            List<T> dataTransferObjects,
            String... sheetNames
    ) {
        if (dataTransferObjects == null || dataTransferObjects.isEmpty()) {
            throw new IllegalArgumentException("DTO list cannot be null or empty");
        }
        var clazz = dataTransferObjects.getFirst().getClass();

        List<Field> fields = Arrays.stream(clazz.getDeclaredFields())
                .filter(f -> f.isAnnotationPresent(ExcelWriteColumn.class))
                .sorted(Comparator.comparingInt(f -> f.getAnnotation(ExcelWriteColumn.class).order()))
                .toList();

        var headers = fields.stream()
                .map(f -> f.getAnnotation(ExcelWriteColumn.class).headerName())
                .collect(Collectors.toList());

        var rows = new ArrayList<List<Object>>();
        for (var dto : dataTransferObjects) {
            var row = new ArrayList<>();
            for (var field : fields) {
                field.setAccessible(true);
                try {
                    row.add(field.get(dto));
                } catch (IllegalAccessException e) {
                    row.add(null);
                }
            }
            rows.add(row);
        }

        var actualSheetName = sheetNames.length > 0 ? sheetNames[0] : clazz.getSimpleName();
        return ExcelDocument.builder()
                .sheet(actualSheetName)
                .headers(headers)
                .rows(rows)
                .build();
    }

    public <T> List<T> toDTO(
            Class<T> clazz,
            ConverterPort converter
    ) {
        var result = new ArrayList<T>();
        var headers = this.getHeaders();

        var fields = new ArrayList<Field>();
        for (var field : clazz.getDeclaredFields()) {
            if (field.isAnnotationPresent(ExcelReadColumn.class)) {
                fields.add(field);
            }
        }
        if (fields.isEmpty()) {
            throw new IllegalArgumentException(clazz.getSimpleName() + " has no @NinjaExcelRead fields");
        }

        var fieldToColIdx = new ArrayList<Integer>();
        for (var field : fields) {
            var annotation = field.getAnnotation(ExcelReadColumn.class);
            String headerName = annotation.headerName();
            int idx = headers.indexOf(headerName);
            if (idx == -1) {
                throw new IllegalArgumentException("Header [" + headerName + "] not found in Excel headers");
            }
            fieldToColIdx.add(idx);
        }

        for (var row : this.getRows()) {
            try {
                T dto = clazz.getDeclaredConstructor().newInstance();
                for (int i = 0; i < fields.size(); i++) {
                    var field = fields.get(i);
                    var colIdx = fieldToColIdx.get(i);

                    var cellValue = row.get(colIdx);
                    // 어노테이션에 type, defaultValue 있으면 우선 적용
                    var annotation = field.getAnnotation(ExcelReadColumn.class);
                    var targetType = annotation.type() == Void.class ? field.getType() : annotation.type();
                    var valueToSet = cellValue != null
                            ? converter.convert(cellValue, targetType)
                            : annotation.defaultValue().isEmpty() ? null : converter.convert(annotation.defaultValue(), targetType);

                    field.setAccessible(true);
                    field.set(dto, valueToSet);
                }
                result.add(dto);
            } catch (Exception e) {
                throw new RuntimeException("Failed to map row to DTO", e);
            }
        }
        return result;
    }
}
