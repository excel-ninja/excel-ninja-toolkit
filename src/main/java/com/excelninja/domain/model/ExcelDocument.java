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
        this.headers = Collections.unmodifiableList(new ArrayList<>(Objects.requireNonNull(headers)));
        this.rows = Collections.unmodifiableList(new ArrayList<>(Objects.requireNonNull(rows)));
        this.columnWidths = Collections.unmodifiableMap(new HashMap<>(Objects.requireNonNull(columnWidths)));
        this.rowHeights = Collections.unmodifiableMap(new HashMap<>(Objects.requireNonNull(rowHeights)));
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

        public Builder headers(List<String> headerList) {
            Objects.requireNonNull(headerList);
            this.headers.clear();
            this.headers.addAll(headerList);
            return this;
        }

        public Builder rows(List<List<Object>> rowValuesList) {
            for (List<Object> rowValues : rowValuesList) {
                Objects.requireNonNull(rowValues);
                if (rowValues.size() != headers.size()) {
                    throw new IllegalArgumentException("Row values size (" + rowValues.size() + ") does not match header count (" + headers.size() + ")");
                }
                this.rows.add(new ArrayList<>(rowValues));
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
        Class<?> dtoClass = dataTransferObjects.get(0).getClass();

        List<Field> annotatedFields = Arrays.stream(dtoClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelWriteColumn.class))
                .sorted(Comparator.comparingInt(field -> field.getAnnotation(ExcelWriteColumn.class).order()))
                .collect(Collectors.toList());

        List<String> headerList = annotatedFields.stream()
                .map(field -> field.getAnnotation(ExcelWriteColumn.class).headerName())
                .collect(Collectors.toList());

        List<List<Object>> rowsList = new ArrayList<>();
        for (T dtoInstance : dataTransferObjects) {
            List<Object> rowList = new ArrayList<>();
            for (Field field : annotatedFields) {
                field.setAccessible(true);
                try {
                    rowList.add(field.get(dtoInstance));
                } catch (IllegalAccessException exception) {
                    rowList.add(null);
                }
            }
            rowsList.add(rowList);
        }

        String actualSheetName = sheetNames.length > 0 ? sheetNames[0] : dtoClass.getSimpleName();
        return ExcelDocument.builder()
                .sheet(actualSheetName)
                .headers(headerList)
                .rows(rowsList)
                .build();
    }

    public <T> List<T> toDTO(
            Class<T> dtoClass,
            ConverterPort converter
    ) {
        List<T> resultList = new ArrayList<T>();
        List<String> excelHeaders = this.getHeaders();

        List<Field> annotatedFields = new ArrayList<Field>();
        for (Field field : dtoClass.getDeclaredFields()) {
            if (field.isAnnotationPresent(ExcelReadColumn.class)) {
                annotatedFields.add(field);
            }
        }

        if (annotatedFields.isEmpty()) {
            throw new IllegalArgumentException(dtoClass.getSimpleName() + " has no @ExcelReadColumn fields");
        }

        List<Integer> fieldToColumnIndex = new ArrayList<Integer>();
        for (Field field : annotatedFields) {
            ExcelReadColumn annotation = field.getAnnotation(ExcelReadColumn.class);
            String headerName = annotation.headerName();
            int headerIndex = excelHeaders.indexOf(headerName);
            if (headerIndex == -1) {
                throw new IllegalArgumentException("Header [" + headerName + "] not found in Excel headers");
            }
            fieldToColumnIndex.add(headerIndex);
        }

        for (List<Object> rowData : this.getRows()) {
            try {
                T dtoInstance = dtoClass.getDeclaredConstructor().newInstance();
                for (int i = 0; i < annotatedFields.size(); i++) {
                    Field field = annotatedFields.get(i);
                    int columnIndex = fieldToColumnIndex.get(i);

                    Object cellValue = rowData.get(columnIndex);
                    ExcelReadColumn annotation = field.getAnnotation(ExcelReadColumn.class);
                    Class<?> targetType = annotation.type() == Void.class ? field.getType() : annotation.type();
                    Object valueToSet = null;
                    if (cellValue != null) {
                        valueToSet = converter.convert(cellValue, targetType);
                    } else if (!annotation.defaultValue().isEmpty()) {
                        valueToSet = converter.convert(annotation.defaultValue(), targetType);
                    }

                    field.setAccessible(true);
                    field.set(dtoInstance, valueToSet);
                }
                resultList.add(dtoInstance);
            } catch (Exception exception) {
                throw new RuntimeException("Failed to map row to DTO", exception);
            }
        }
        return resultList;
    }
}
