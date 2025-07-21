package com.excelninja.domain.model;

import com.excelninja.domain.exception.DocumentConversionException;
import com.excelninja.domain.exception.EntityMappingException;
import com.excelninja.infrastructure.metadata.EntityMetadata;
import com.excelninja.infrastructure.metadata.FieldMapping;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class ExcelSheet {
    private final SheetName name;
    private final Headers headers;
    private final DocumentRows rows;
    private final SheetMetadata metadata;

    private ExcelSheet(
            SheetName name,
            Headers headers,
            DocumentRows rows,
            SheetMetadata metadata
    ) {
        this.name = Objects.requireNonNull(name);
        this.headers = Objects.requireNonNull(headers);
        this.rows = Objects.requireNonNull(rows);
        this.metadata = metadata != null ? metadata : new SheetMetadata();
    }

    public static <T> ExcelSheet fromEntities(
            List<T> entities,
            String sheetName
    ) {
        return fromEntities(entities, sheetName, null);
    }

    public static <T> ExcelSheet fromEntities(
            List<T> entities,
            String sheetName,
            SheetMetadata metadata
    ) {
        if (entities == null || entities.isEmpty()) {
            throw new EntityMappingException(Object.class, "Entity list cannot be null or empty");
        }

        Class<?> entityType = entities.get(0).getClass();
        String actualSheetName = sheetName != null ? sheetName : entityType.getSimpleName();

        EntityMetadata<?> entityMetadata = EntityMetadata.of(entityType);
        List<FieldMapping> writeFields = entityMetadata.getWriteFieldMappings();

        if (writeFields.isEmpty()) {
            throw EntityMappingException.noAnnotatedFields(entityType);
        }

        Headers headers = createHeadersFromFields(writeFields);
        DocumentRows rows = createRowsFromEntities(entities, writeFields, headers.size());

        return new ExcelSheet(
                new SheetName(actualSheetName),
                headers,
                rows,
                metadata
        );
    }

    public static SheetBuilder builder() {
        return new SheetBuilder();
    }

    private static Headers createHeadersFromFields(List<FieldMapping> fields) {
        List<Header> headerList = new ArrayList<>();
        for (int i = 0; i < fields.size(); i++) {
            FieldMapping fieldMapping = fields.get(i);
            headerList.add(new Header(fieldMapping.getHeaderName(), i));
        }
        return new Headers(headerList);
    }

    private static DocumentRows createRowsFromEntities(
            List<?> entities,
            List<FieldMapping> fields,
            int expectedColumnCount
    ) {
        List<DocumentRow> documentRows = new ArrayList<>();

        for (int entityIndex = 0; entityIndex < entities.size(); entityIndex++) {
            Object entity = entities.get(entityIndex);
            List<Object> rowValues = new ArrayList<>();

            for (FieldMapping fieldMapping : fields) {
                try {
                    Object value = fieldMapping.getValue(entity);
                    rowValues.add(value);
                } catch (Exception e) {
                    throw new DocumentConversionException("Failed to read field '" + fieldMapping.getFieldName() + "' from entity at index " + entityIndex + ": " + e.getMessage(), e);
                }
            }

            documentRows.add(new DocumentRow(rowValues, entityIndex + 1));
        }

        return new DocumentRows(documentRows, expectedColumnCount);
    }

    public SheetName getName() {
        return name;
    }

    public Headers getHeaders() {
        return headers;
    }

    public DocumentRows getRows() {
        return rows;
    }

    public SheetMetadata getMetadata() {
        return metadata;
    }

    public static class SheetBuilder {
        private SheetName name;
        private Headers headers;
        private DocumentRows rows;
        private SheetMetadata metadata = new SheetMetadata();

        public SheetBuilder name(String name) {
            this.name = new SheetName(name);
            return this;
        }

        public SheetBuilder headers(String... headerNames) {
            this.headers = Headers.of(headerNames);
            return this;
        }

        public SheetBuilder headers(List<String> headerNames) {
            this.headers = Headers.of(headerNames);
            return this;
        }

        public SheetBuilder rows(List<List<Object>> rowData) {
            if (headers == null) {
                throw new IllegalStateException("Headers must be set before rows");
            }
            this.rows = DocumentRows.of(rowData, headers.size());
            return this;
        }

        public SheetBuilder metadata(SheetMetadata metadata) {
            this.metadata = metadata;
            return this;
        }

        public SheetBuilder columnWidth(
                int columnIndex,
                int width
        ) {
            this.metadata = this.metadata.withColumnWidth(columnIndex, width);
            return this;
        }

        public SheetBuilder rowHeight(
                int rowIndex,
                short height
        ) {
            this.metadata = this.metadata.withRowHeight(rowIndex, height);
            return this;
        }

        public ExcelSheet build() {
            if (name == null) {
                name = SheetName.defaultName();
            }
            if (headers == null) {
                throw new IllegalStateException("Headers must be provided");
            }
            if (rows == null) {
                rows = DocumentRows.empty(headers.size());
            }
            return new ExcelSheet(name, headers, rows, metadata);
        }
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
}