package com.excelninja.domain.model;

import com.excelninja.domain.exception.EntityMappingException;

import java.util.*;

public class ExcelWorkbook {
    private final Map<String, ExcelSheet> sheets;
    private final WorkbookMetadata metadata;

    private ExcelWorkbook(
            Map<String, ExcelSheet> sheets,
            WorkbookMetadata metadata
    ) {
        if (sheets == null || sheets.isEmpty()) {
            throw new IllegalArgumentException("At least one sheet is required");
        }
        this.sheets = Collections.unmodifiableMap(new LinkedHashMap<>(sheets));
        this.metadata = metadata != null ? metadata : new WorkbookMetadata();
    }

    public static WorkbookBuilder builder() {
        return new WorkbookBuilder();
    }

    public Set<String> getSheetNames() {
        return sheets.keySet();
    }

    public ExcelSheet getSheet(String name) {
        return sheets.get(name);
    }

    public WorkbookMetadata getMetadata() {
        return metadata;
    }

    public static class WorkbookBuilder {
        private final Map<String, ExcelSheet> sheets = new LinkedHashMap<>();
        private WorkbookMetadata metadata;

        public WorkbookBuilder sheet(
                String name,
                ExcelSheet sheet
        ) {
            if (name == null || name.trim().isEmpty()) {
                throw new IllegalArgumentException("Sheet name cannot be null or empty");
            }
            this.sheets.put(name, sheet);
            return this;
        }

        public <T> WorkbookBuilder sheet(
                String name,
                List<T> entities
        ) {
            if (entities == null || entities.isEmpty()) {
                throw new EntityMappingException(Object.class, "Entity list cannot be null or empty");
            }
            return sheet(name, ExcelSheet.fromEntities(entities, name));
        }

        public <T> WorkbookBuilder sheet(List<T> entities) {
            if (entities == null || entities.isEmpty()) {
                throw new EntityMappingException(Object.class, "Entity list cannot be null or empty");
            }
            String sheetName = entities.get(0).getClass().getSimpleName();
            return sheet(sheetName, ExcelSheet.fromEntities(entities, sheetName));
        }

        public WorkbookBuilder metadata(WorkbookMetadata metadata) {
            this.metadata = metadata;
            return this;
        }

        public ExcelWorkbook build() {
            return new ExcelWorkbook(sheets, metadata);
        }
    }
}