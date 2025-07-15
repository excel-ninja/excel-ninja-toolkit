package com.excelninja.domain.exception;

public class InvalidDocumentStructureException extends ExcelDomainException {

    public InvalidDocumentStructureException(String message) {
        super(message);
    }

    public static InvalidDocumentStructureException emptyHeaders() {
        return new InvalidDocumentStructureException("Excel document must have at least one header");
    }

    public static InvalidDocumentStructureException rowColumnMismatch(
            int expectedColumns,
            int actualColumns,
            int rowIndex
    ) {
        return new InvalidDocumentStructureException(
                String.format("Row %d has %d columns but expected %d", rowIndex, actualColumns, expectedColumns)
        );
    }

    public static InvalidDocumentStructureException emptySheetName() {
        return new InvalidDocumentStructureException("Sheet name cannot be null or empty");
    }
}