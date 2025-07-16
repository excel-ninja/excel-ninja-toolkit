package com.excelninja.domain.model;

import com.excelninja.domain.exception.InvalidDocumentStructureException;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

public final class DocumentRows {

    private final List<DocumentRow> rows;
    private final int expectedColumnCount;

    public DocumentRows(
            List<DocumentRow> rows,
            int expectedColumnCount
    ) {
        validateRows(rows, expectedColumnCount);
        this.rows = Collections.unmodifiableList(new ArrayList<>(rows));
        this.expectedColumnCount = expectedColumnCount;
    }

    private void validateRows(
            List<DocumentRow> rows,
            int expectedColumnCount
    ) {
        if (rows == null) {
            throw new InvalidDocumentStructureException("Rows cannot be null");
        }

        if (expectedColumnCount < 0) {
            throw new InvalidDocumentStructureException("Expected column count cannot be negative");
        }

        for (DocumentRow row : rows) {
            if (row == null) {
                throw new InvalidDocumentStructureException("Row cannot be null");
            }

            if (row.getColumnCount() != expectedColumnCount) {
                throw InvalidDocumentStructureException.rowColumnMismatch(expectedColumnCount, row.getColumnCount(), row.getRowNumber());
            }
        }
    }

    public List<DocumentRow> getRows() {
        return rows;
    }

    public int size() {
        return rows.size();
    }

    public boolean isEmpty() {
        return rows.isEmpty();
    }

    public DocumentRow getRow(int index) {
        if (index < 0 || index >= rows.size()) {
            throw new InvalidDocumentStructureException(
                    String.format("Row index %d is out of bounds. Document has %d rows",
                            index, rows.size()));
        }
        return rows.get(index);
    }

    public List<Object> getColumn(int columnIndex) {
        if (columnIndex < 0 || columnIndex >= expectedColumnCount) {
            throw new InvalidDocumentStructureException(
                    String.format("Column index %d is out of bounds. Expected column count: %d",
                            columnIndex, expectedColumnCount));
        }

        return rows.stream()
                .map(row -> row.getValue(columnIndex))
                .collect(Collectors.toList());
    }

    public List<Object> getColumnByHeader(
            Headers headers,
            String headerName
    ) {
        int position = headers.getPositionOf(headerName);
        return getColumn(position);
    }

    public int getExpectedColumnCount() {
        return expectedColumnCount;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        DocumentRows that = (DocumentRows) o;
        return expectedColumnCount == that.expectedColumnCount && Objects.equals(rows, that.rows);
    }

    @Override
    public int hashCode() {
        return Objects.hash(rows, expectedColumnCount);
    }

    @Override
    public String toString() {
        return String.format("DocumentRows{rowCount=%d, columnCount=%d}",
                rows.size(), expectedColumnCount);
    }

    public static DocumentRows of(
            List<List<Object>> rawRows,
            int expectedColumnCount
    ) {
        List<DocumentRow> documentRows = new ArrayList<>();
        for (int i = 0; i < rawRows.size(); i++) {
            documentRows.add(new DocumentRow(rawRows.get(i), i + 1));
        }
        return new DocumentRows(documentRows, expectedColumnCount);
    }

    public static DocumentRows empty(int expectedColumnCount) {
        return new DocumentRows(Collections.emptyList(), expectedColumnCount);
    }
}