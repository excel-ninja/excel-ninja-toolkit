package com.excelninja.domain.model;

import com.excelninja.domain.annotation.ExcelColumn;

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

    public static Builder builder(String sheetName) {
        return new Builder(sheetName);
    }

    public static final class Builder {
        private final String sheetName;
        private final List<String> headers = new ArrayList<>();
        private final List<List<Object>> rows = new ArrayList<>();
        private final Map<Integer, Integer> columnWidths = new HashMap<>();
        private final Map<Integer, Short> rowHeights = new HashMap<>();

        public Builder(String sheetName) {
            this.sheetName = sheetName;
        }

        public Builder withHeaders(List<String> headers) {
            Objects.requireNonNull(headers);
            this.headers.clear();
            this.headers.addAll(headers);
            return this;
        }

        public Builder addRow(List<Object> rowValues) {
            Objects.requireNonNull(rowValues);
            if (rowValues.size() != headers.size()) {
                throw new IllegalArgumentException(
                        "Row values size (" + rowValues.size() +
                                ") does not match header count (" + headers.size() + ")"
                );
            }
            this.rows.add(new ArrayList<>(rowValues));
            return this;
        }

        public Builder columnWidth(
                int columnIndex,
                int width
        ) {
            columnWidths.put(columnIndex, width);
            return this;
        }

        public Builder rowHeight(
                int rowIndex,
                short height
        ) {
            rowHeights.put(rowIndex, height);
            return this;
        }

        public ExcelDocument build() {
            return new ExcelDocument(sheetName, headers, rows, columnWidths, rowHeights);
        }
    }

    public static <T> ExcelDocument fromDtoList(
            List<T> dtos,
            String... sheetNames
    ) {
        if (dtos == null || dtos.isEmpty()) {
            throw new IllegalArgumentException("DTO list cannot be null or empty");
        }
        var clazz = dtos.get(0).getClass();

        List<Field> fields = Arrays.stream(clazz.getDeclaredFields())
                .filter(f -> f.isAnnotationPresent(ExcelColumn.class))
                .sorted(Comparator.comparingInt(f -> f.getAnnotation(ExcelColumn.class).order()))
                .toList();

        var headers = fields.stream()
                .map(f -> f.getAnnotation(ExcelColumn.class).headerName())
                .collect(Collectors.toList());

        var rows = new ArrayList<List<Object>>();
        for (var dto : dtos) {
            var row = new ArrayList<Object>();
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
        var builder = ExcelDocument.builder(actualSheetName).withHeaders(headers);
        for (var row : rows) {
            builder.addRow(row);
        }
        return builder.build();
    }
}
