package com.excelninja.domain.model;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

public class SheetMetadata {
    private final Map<Integer, Integer> columnWidths;
    private final Map<Integer, Short> rowHeights;
    private final boolean autoSizeColumns;

    public SheetMetadata() {
        this(new HashMap<>(), new HashMap<>(), true);
    }

    public SheetMetadata(
            Map<Integer, Integer> columnWidths,
            Map<Integer, Short> rowHeights,
            boolean autoSizeColumns
    ) {
        this.columnWidths = new HashMap<>(columnWidths);
        this.rowHeights = new HashMap<>(rowHeights);
        this.autoSizeColumns = autoSizeColumns;
    }

    public Map<Integer, Integer> getColumnWidths() {
        return Collections.unmodifiableMap(columnWidths);
    }

    public Map<Integer, Short> getRowHeights() {
        return Collections.unmodifiableMap(rowHeights);
    }

    public boolean isAutoSizeColumns() {
        return autoSizeColumns;
    }

    public SheetMetadata withColumnWidth(
            int columnIndex,
            int width
    ) {
        Map<Integer, Integer> newColumnWidths = new HashMap<>(this.columnWidths);
        newColumnWidths.put(columnIndex, width);
        return new SheetMetadata(newColumnWidths, rowHeights, autoSizeColumns);
    }

    public SheetMetadata withRowHeight(
            int rowIndex,
            short height
    ) {
        Map<Integer, Short> newRowHeights = new HashMap<>(this.rowHeights);
        newRowHeights.put(rowIndex, height);
        return new SheetMetadata(columnWidths, newRowHeights, autoSizeColumns);
    }
}