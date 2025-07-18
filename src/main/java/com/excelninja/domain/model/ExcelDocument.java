package com.excelninja.domain.model;

import com.excelninja.application.port.ConverterPort;
import com.excelninja.domain.annotation.ExcelWriteColumn;
import com.excelninja.domain.exception.DocumentConversionException;
import com.excelninja.domain.exception.EntityMappingException;
import com.excelninja.domain.exception.HeaderMismatchException;
import com.excelninja.domain.exception.InvalidDocumentStructureException;
import com.excelninja.infrastructure.metadata.EntityMetadata;
import com.excelninja.infrastructure.metadata.FieldMapping;
import com.excelninja.infrastructure.util.ReflectionUtils;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelDocument {

    private final SheetName sheetName;
    private final Headers headers;
    private final DocumentRows rows;
    private final Map<Integer, Integer> columnWidths;
    private final Map<Integer, Short> rowHeights;

    private ExcelDocument(
            SheetName sheetName,
            Headers headers,
            DocumentRows rows,
            Map<Integer, Integer> columnWidths,
            Map<Integer, Short> rowHeights
    ) {
        this.sheetName = Objects.requireNonNull(sheetName, "Sheet name cannot be null");
        this.headers = Objects.requireNonNull(headers, "Headers cannot be null");
        this.rows = Objects.requireNonNull(rows, "Rows cannot be null");
        this.columnWidths = Collections.unmodifiableMap(new HashMap<>(columnWidths != null ? columnWidths : new HashMap<>()));
        this.rowHeights = Collections.unmodifiableMap(new HashMap<>(rowHeights != null ? rowHeights : new HashMap<>()));

        if (rows.getExpectedColumnCount() != headers.size()) {
            throw InvalidDocumentStructureException.rowColumnMismatch(headers.size(), rows.getExpectedColumnCount(), 0);
        }
    }

    public SheetName getSheetName() {
        return sheetName;
    }

    public Headers getHeaders() {
        return headers;
    }

    public DocumentRows getRows() {
        return rows;
    }

    public Map<Integer, Integer> getColumnWidths() {
        return columnWidths;
    }

    public Map<Integer, Short> getRowHeights() {
        return rowHeights;
    }

    public Object getCellValue(
            int rowIndex,
            String headerName
    ) {
        return rows.getRow(rowIndex).getValueByHeader(headers, headerName);
    }

    public <T> T getCellValue(
            int rowIndex,
            String headerName,
            Class<T> type
    ) {
        return rows.getRow(rowIndex).getValueByHeader(headers, headerName, type);
    }

    public List<Object> getColumn(String headerName) {
        return rows.getColumnByHeader(headers, headerName);
    }

    public boolean hasData() {
        return !rows.isEmpty();
    }

    public int getRowCount() {
        return rows.size();
    }

    public int getColumnCount() {
        return headers.size();
    }

    public <T> List<T> convertToEntities(
            Class<T> entityType,
            ConverterPort converter
    ) {
        EntityMetadata<T> metadata = EntityMetadata.of(entityType);

        List<Integer> fieldToColumnMapping = createFieldToColumnMappingOptimized(metadata);

        return convertRowsToEntitiesOptimized(entityType, converter, metadata, fieldToColumnMapping);
    }

    private List<Integer> createFieldToColumnMappingOptimized(EntityMetadata<?> metadata) {
        List<Integer> mapping = new ArrayList<>();

        for (FieldMapping fieldMapping : metadata.getReadFieldMappings()) {
            String headerName = fieldMapping.getHeaderName();

            if (!headers.containsHeader(headerName)) {
                throw HeaderMismatchException.headerNotFound(headerName);
            }

            mapping.add(headers.getPositionOf(headerName));
        }

        return mapping;
    }

    private <T> List<T> convertRowsToEntitiesOptimized(
            Class<T> entityType,
            ConverterPort converter,
            EntityMetadata<T> metadata,
            List<Integer> fieldToColumnMapping
    ) {
        List<T> entities = new ArrayList<>();
        List<FieldMapping> fieldMappings = metadata.getReadFieldMappings();

        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            DocumentRow row = rows.getRow(rowIndex);

            try {
                T entity = metadata.createInstance();

                for (int fieldIndex = 0; fieldIndex < fieldMappings.size(); fieldIndex++) {
                    FieldMapping fieldMapping = fieldMappings.get(fieldIndex);
                    int columnIndex = fieldToColumnMapping.get(fieldIndex);

                    Object cellValue = row.getValue(columnIndex);

                    fieldMapping.setValue(entity, cellValue, converter);
                }

                entities.add(entity);

            } catch (Exception e) {
                throw new DocumentConversionException("Failed to create entity of type " + entityType.getName() + " for row " + (rowIndex + 1), e);
            }
        }

        return entities;
    }

    public static DocumentReader reader() {
        return new DocumentReader();
    }

    public static final class DocumentReader {
        private SheetName sheetName;
        private Headers headers;
        private DocumentRows rows;
        private final Map<Integer, Integer> columnWidths = new HashMap<>();
        private final Map<Integer, Short> rowHeights = new HashMap<>();

        public DocumentReader() {}

        public DocumentReader sheet(String sheetName) {
            this.sheetName = new SheetName(sheetName);
            return this;
        }

        public DocumentReader sheet(SheetName sheetName) {
            this.sheetName = Objects.requireNonNull(sheetName);
            return this;
        }

        public DocumentReader headers(List<String> headerNames) {
            this.headers = Headers.of(headerNames);
            return this;
        }

        public DocumentReader headers(Headers headers) {
            this.headers = Objects.requireNonNull(headers);
            return this;
        }

        public DocumentReader headers(String... headerNames) {
            this.headers = Headers.of(headerNames);
            return this;
        }

        public DocumentReader rows(List<List<Object>> rowValuesList) {
            if (headers == null) {
                throw new IllegalStateException("Headers must be set before rows");
            }
            this.rows = DocumentRows.of(rowValuesList, headers.size());
            return this;
        }

        public DocumentReader rows(DocumentRows rows) {
            this.rows = Objects.requireNonNull(rows);
            return this;
        }

        public DocumentReader columnWidth(
                int columnIndex,
                int width
        ) {
            this.columnWidths.put(columnIndex, width);
            return this;
        }

        public DocumentReader rowHeight(
                int rowIndex,
                short height
        ) {
            this.rowHeights.put(rowIndex, height);
            return this;
        }

        public ExcelDocument create() {
            if (sheetName == null) {
                sheetName = SheetName.defaultName();
            }
            if (headers == null) {
                throw new IllegalStateException("Headers must be provided");
            }
            if (rows == null) {
                rows = DocumentRows.empty(headers.size());
            }

            return new ExcelDocument(sheetName, headers, rows, columnWidths, rowHeights);
        }
    }

    public static Writer writer() {
        return new Writer();
    }

    public static final class Writer {
        public <T> DocumentWriter<T> objects(List<T> objects) {
            return new DocumentWriter<>(objects);
        }
    }

    public static final class DocumentWriter<T> {
        private final List<T> objects;
        private String sheetName;
        private final Map<Integer, Integer> columnWidths = new HashMap<>();
        private final Map<Integer, Short> rowHeights = new HashMap<>();

        private DocumentWriter(List<T> objects) {
            if (objects == null || objects.isEmpty()) {
                throw EntityMappingException.emptyEntityList();
            }
            this.objects = objects;
        }

        public DocumentWriter<T> sheetName(String sheetName) {
            this.sheetName = sheetName;
            return this;
        }

        public DocumentWriter<T> columnWidth(
                int columnIndex,
                int width
        ) {
            this.columnWidths.put(columnIndex, width);
            return this;
        }

        public DocumentWriter<T> rowHeight(
                int rowIndex,
                short height
        ) {
            this.rowHeights.put(rowIndex, height);
            return this;
        }

        public ExcelDocument create() {
            Class<?> entityType = objects.get(0).getClass();

            List<Field> writeFields = extractAndValidateWriteFieldsLegacy(entityType);
            Headers headers = createHeadersFromFieldsLegacy(writeFields);
            DocumentRows rows = createRowsFromEntitiesLegacy(objects, writeFields, headers.size());

            SheetName sheet = sheetName != null ? new SheetName(sheetName) : new SheetName(entityType.getSimpleName());
            return new ExcelDocument(sheet, headers, rows, columnWidths, rowHeights);
        }
    }

    private static List<Field> extractAndValidateWriteFieldsLegacy(Class<?> entityType) {
        List<Field> annotatedFields = Arrays.stream(entityType.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelWriteColumn.class))
                .collect(Collectors.toList());

        if (annotatedFields.isEmpty()) {
            throw EntityMappingException.noAnnotatedFields(entityType);
        }

        validateWriteAnnotationsLegacy(annotatedFields, entityType);
        annotatedFields.sort(Comparator.comparingInt(field -> field.getAnnotation(ExcelWriteColumn.class).order()));

        return annotatedFields;
    }

    private static void validateWriteAnnotationsLegacy(
            List<Field> fields,
            Class<?> entityType
    ) {
        Set<String> usedHeaders = new HashSet<>();
        Set<Integer> usedOrders = new HashSet<>();

        for (Field field : fields) {
            ExcelWriteColumn annotation = field.getAnnotation(ExcelWriteColumn.class);

            String headerName = annotation.headerName();
            if (headerName == null || headerName.trim().isEmpty()) {
                throw EntityMappingException.invalidAnnotationConfiguration(entityType, "Empty header name in field: " + field.getName());
            }

            if (!usedHeaders.add(headerName)) {
                throw EntityMappingException.invalidAnnotationConfiguration(entityType, "Duplicate header name: " + headerName);
            }

            int order = annotation.order();
            if (order >= 0 && order != Integer.MAX_VALUE && !usedOrders.add(order)) {
                throw EntityMappingException.invalidAnnotationConfiguration(entityType, "Duplicate order value: " + order);
            }
        }
    }

    private static Headers createHeadersFromFieldsLegacy(List<Field> fields) {
        List<Header> headerList = new ArrayList<>();
        for (int i = 0; i < fields.size(); i++) {
            Field field = fields.get(i);
            ExcelWriteColumn annotation = field.getAnnotation(ExcelWriteColumn.class);
            headerList.add(new Header(annotation.headerName(), i));
        }
        return new Headers(headerList);
    }

    private static DocumentRows createRowsFromEntitiesLegacy(
            List<?> entities,
            List<Field> fields,
            int expectedColumnCount
    ) {
        List<DocumentRow> documentRows = new ArrayList<>();

        for (int entityIndex = 0; entityIndex < entities.size(); entityIndex++) {
            Object entity = entities.get(entityIndex);
            List<Object> rowValues = new ArrayList<>();

            for (Field field : fields) {
                try {
                    Object value = ReflectionUtils.getFieldValue(entity, field);
                    rowValues.add(value);
                } catch (DocumentConversionException e) {
                    throw new DocumentConversionException("Failed to read field '" + field.getName() + "' from entity at index " + entityIndex + ": " + e.getMessage(), e);
                }
            }

            documentRows.add(new DocumentRow(rowValues, entityIndex + 1));
        }

        return new DocumentRows(documentRows, expectedColumnCount);
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        ExcelDocument that = (ExcelDocument) o;
        return Objects.equals(sheetName, that.sheetName) && Objects.equals(headers, that.headers) && Objects.equals(rows, that.rows);
    }

    @Override
    public int hashCode() {
        return Objects.hash(sheetName, headers, rows);
    }

    @Override
    public String toString() {
        return String.format("ExcelDocument{sheetName=%s, headers=%s, rowCount=%d}", sheetName, headers, rows.size());
    }
}