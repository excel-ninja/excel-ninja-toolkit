package com.excelninja.domain.model;

import com.excelninja.application.port.ConverterPort;
import com.excelninja.domain.annotation.ExcelReadColumn;
import com.excelninja.domain.annotation.ExcelWriteColumn;
import com.excelninja.domain.exception.*;
import com.excelninja.infrastructure.util.ReflectionUtils;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
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
        if (sheetName == null || sheetName.trim().isEmpty()) {
            throw InvalidDocumentStructureException.emptySheetName();
        }

        if (headers == null || headers.isEmpty()) {
            throw InvalidDocumentStructureException.emptyHeaders();
        }

        if (rows != null) {
            for (int i = 0; i < rows.size(); i++) {
                List<Object> row = rows.get(i);
                if (row != null && row.size() != headers.size()) {
                    throw InvalidDocumentStructureException.rowColumnMismatch(
                            headers.size(), row.size(), i + 1);
                }
            }
        }

        this.sheetName = sheetName;
        this.headers = Collections.unmodifiableList(new ArrayList<>(headers));
        this.rows = Collections.unmodifiableList(new ArrayList<>(rows != null ? rows : new ArrayList<>()));
        this.columnWidths = Collections.unmodifiableMap(new HashMap<>(columnWidths != null ? columnWidths : new HashMap<>()));
        this.rowHeights = Collections.unmodifiableMap(new HashMap<>(rowHeights != null ? rowHeights : new HashMap<>()));
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
            Objects.requireNonNull(headerList, "Headers cannot be null");

            if (headerList.isEmpty()) {
                throw InvalidDocumentStructureException.emptyHeaders();
            }

            Set<String> uniqueHeaders = new HashSet<>();
            for (String header : headerList) {
                if (header == null || header.trim().isEmpty()) {
                    throw InvalidDocumentStructureException.emptySheetName();
                }
                if (!uniqueHeaders.add(header.trim())) {
                    throw HeaderMismatchException.duplicateHeader(header);
                }
            }

            this.headers.clear();
            this.headers.addAll(headerList);
            return this;
        }

        public Builder rows(List<List<Object>> rowValuesList) {
            Objects.requireNonNull(rowValuesList, "Rows cannot be null");

            for (int i = 0; i < rowValuesList.size(); i++) {
                List<Object> rowValues = rowValuesList.get(i);
                Objects.requireNonNull(rowValues, "Row values cannot be null");

                if (rowValues.size() != headers.size()) {
                    throw InvalidDocumentStructureException.rowColumnMismatch(
                            headers.size(), rowValues.size(), i + 1);
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
            throw EntityMappingException.emptyEntityList();
        }

        Class<?> dtoClass = dataTransferObjects.get(0).getClass();

        List<Field> annotatedFields = Arrays.stream(dtoClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelWriteColumn.class))
                .collect(Collectors.toList());

        if (annotatedFields.isEmpty()) {
            throw EntityMappingException.noAnnotatedFields(dtoClass);
        }

        validateWriteAnnotations(annotatedFields, dtoClass);

        annotatedFields.sort(Comparator.comparingInt(field -> field.getAnnotation(ExcelWriteColumn.class).order()));

        List<String> headerList = annotatedFields.stream()
                .map(field -> field.getAnnotation(ExcelWriteColumn.class).headerName())
                .collect(Collectors.toList());

        List<List<Object>> rowsList = new ArrayList<>();
        for (T dtoInstance : dataTransferObjects) {
            List<Object> rowList = new ArrayList<>();
            for (Field field : annotatedFields) {
                try {
                    Object value = ReflectionUtils.getFieldValue(dtoInstance, field);
                    rowList.add(value);
                } catch (DocumentConversionException e) {
                    throw new DocumentConversionException("Failed to read field '" + field.getName() + "' from " + dtoClass.getSimpleName() + ": " + e.getMessage(), e);
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
        List<T> resultList = new ArrayList<>();
        List<String> excelHeaders = this.getHeaders();

        List<Field> annotatedFields = Arrays.stream(dtoClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelReadColumn.class))
                .collect(Collectors.toList());

        if (annotatedFields.isEmpty()) {
            throw EntityMappingException.noAnnotatedFields(dtoClass);
        }

        List<Integer> fieldToColumnIndex = validateAndMapHeaders(annotatedFields, excelHeaders, dtoClass);

        for (int rowIdx = 0; rowIdx < this.getRows().size(); rowIdx++) {
            List<Object> rowData = this.getRows().get(rowIdx);
            try {
                T dtoInstance = dtoClass.getDeclaredConstructor().newInstance();

                for (int i = 0; i < annotatedFields.size(); i++) {
                    Field field = annotatedFields.get(i);
                    int columnIndex = fieldToColumnIndex.get(i);

                    Object cellValue = rowData.get(columnIndex);
                    ExcelReadColumn annotation = field.getAnnotation(ExcelReadColumn.class);
                    Class<?> targetType = annotation.type() == Void.class ? field.getType() : annotation.type();

                    Object valueToSet = null;
                    try {
                        if (cellValue != null) {
                            valueToSet = converter.convert(cellValue, targetType);
                        } else if (!annotation.defaultValue().isEmpty()) {
                            valueToSet = converter.convert(annotation.defaultValue(), targetType);
                        }
                    } catch (Exception conversionException) {
                        throw new DocumentConversionException(field.getName(), cellValue, "Type conversion failed: " + conversionException.getMessage());
                    }

                    try {
                        ReflectionUtils.setFieldValue(dtoInstance, field, valueToSet);
                    } catch (DocumentConversionException e) {
                        throw new DocumentConversionException("Failed to set field '" + field.getName() + "' in " + dtoClass.getSimpleName() + " at row " + (rowIdx + 1) + ": " + e.getMessage(), e);
                    }
                }
                resultList.add(dtoInstance);

            } catch (NoSuchMethodException e) {
                throw new DocumentConversionException("No default constructor found for " + dtoClass.getName() + ". Please add a public no-argument constructor.", e);
            } catch (InstantiationException e) {
                throw new DocumentConversionException("Cannot instantiate " + dtoClass.getName() + " (abstract class or interface). Please use a concrete class.", e);
            } catch (IllegalAccessException e) {
                throw new DocumentConversionException("Cannot access constructor of " + dtoClass.getName() + ". Please make the constructor public.", e);
            } catch (InvocationTargetException e) {
                throw new DocumentConversionException("Constructor threw exception for " + dtoClass.getName() + ": " + e.getCause().getMessage(), e.getCause());
            } catch (ExcelDomainException e) {
                throw e;
            } catch (Exception e) {
                throw new DocumentConversionException("Failed to create instance for row " + (rowIdx + 1) + " of type " + dtoClass.getName(), e);
            }
        }
        return resultList;
    }

    private static void validateWriteAnnotations(List<Field> annotatedFields, Class<?> dtoClass) {
        Set<Integer> usedOrders = new HashSet<>();
        Set<String> usedHeaders = new HashSet<>();

        for (Field field : annotatedFields) {
            ExcelWriteColumn annotation = field.getAnnotation(ExcelWriteColumn.class);

            String headerName = annotation.headerName();
            if (headerName == null || headerName.trim().isEmpty()) {
                throw EntityMappingException.invalidAnnotationConfiguration(dtoClass, "Empty header name in field: " + field.getName());
            }

            if (!usedHeaders.add(headerName)) {
                throw EntityMappingException.invalidAnnotationConfiguration(dtoClass, "Duplicate header name: " + headerName);
            }

            int order = annotation.order();
            if (order >= 0 && order != Integer.MAX_VALUE && !usedOrders.add(order)) {
                throw EntityMappingException.invalidAnnotationConfiguration(dtoClass, "Duplicate order value: " + order);
            }
        }
    }

    private<T> List<Integer> validateAndMapHeaders(List<Field> annotatedFields, List<String> excelHeaders, Class<T> dtoClass) {
        List<Integer> fieldToColumnIndex = new ArrayList<>();

        for (Field field : annotatedFields) {
            ExcelReadColumn annotation = field.getAnnotation(ExcelReadColumn.class);
            String headerName = annotation.headerName();

            if (headerName == null || headerName.trim().isEmpty()) {
                throw EntityMappingException.invalidAnnotationConfiguration(dtoClass, "Empty header name in field: " + field.getName());
            }

            int headerIndex = excelHeaders.indexOf(headerName);
            if (headerIndex == -1) {
                throw HeaderMismatchException.headerNotFound(headerName);
            }
            fieldToColumnIndex.add(headerIndex);
        }

        return fieldToColumnIndex;
    }
}
