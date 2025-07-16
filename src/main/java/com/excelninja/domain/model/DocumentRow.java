package com.excelninja.domain.model;

import com.excelninja.domain.exception.InvalidDocumentStructureException;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;

public final class DocumentRow {

    private final List<Object> values;
    private final int rowNumber;

    public DocumentRow(
            List<Object> values,
            int rowNumber
    ) {
        validateRow(values, rowNumber);
        this.values = Collections.unmodifiableList(new ArrayList<>(values));
        this.rowNumber = rowNumber;
    }

    private void validateRow(
            List<Object> values,
            int rowNumber
    ) {
        if (values == null) {
            throw new InvalidDocumentStructureException("Row values cannot be null");
        }

        if (rowNumber < 0) {
            throw new InvalidDocumentStructureException("Row number cannot be negative: " + rowNumber);
        }
    }

    public Object getValue(int columnIndex) {
        if (columnIndex < 0 || columnIndex >= values.size()) {
            throw new InvalidDocumentStructureException(
                    String.format("Column index %d is out of bounds. Row has %d columns",
                            columnIndex, values.size()));
        }
        return values.get(columnIndex);
    }

    public <T> T getValue(
            int columnIndex,
            Class<T> type
    ) {
        Object value = getValue(columnIndex);
        if (value == null) {
            return null;
        }

        if (type.isInstance(value)) {
            return type.cast(value);
        }

        throw new InvalidDocumentStructureException(
                String.format("Value at column %d is not of type %s. Actual type: %s", columnIndex, type.getSimpleName(), value.getClass().getSimpleName()));
    }

    public Object getValueByHeader(
            Headers headers,
            String headerName
    ) {
        int position = headers.getPositionOf(headerName);
        return getValue(position);
    }

    public <T> T getValueByHeader(
            Headers headers,
            String headerName,
            Class<T> type
    ) {
        int position = headers.getPositionOf(headerName);
        return getValue(position, type);
    }

    public List<Object> getValues() {
        return values;
    }

    public int getColumnCount() {
        return values.size();
    }

    public int getRowNumber() {
        return rowNumber;
    }

    public boolean isEmpty() {
        return values.stream().allMatch(Objects::isNull);
    }

    public boolean hasValue(int columnIndex) {
        return columnIndex >= 0 && columnIndex < values.size() && values.get(columnIndex) != null;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        DocumentRow that = (DocumentRow) o;
        return rowNumber == that.rowNumber && Objects.equals(values, that.values);
    }

    @Override
    public int hashCode() {
        return Objects.hash(values, rowNumber);
    }

    @Override
    public String toString() {
        return String.format("DocumentRow{rowNumber=%d, values=%s}", rowNumber, values);
    }

    public static DocumentRow of(
            List<Object> values,
            int rowNumber
    ) {
        return new DocumentRow(values, rowNumber);
    }
}
