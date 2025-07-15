package com.excelninja.domain.exception;

public class HeaderMismatchException extends ExcelDomainException {

    private final String expectedHeader;
    private final String actualHeader;

    public HeaderMismatchException(
            String expectedHeader,
            String actualHeader
    ) {
        super(String.format("Expected header '%s' but found '%s'", expectedHeader, actualHeader));
        this.expectedHeader = expectedHeader;
        this.actualHeader = actualHeader;
    }

    public static HeaderMismatchException headerNotFound(String requiredHeader) {
        return new HeaderMismatchException(requiredHeader, "not found");
    }

    public static HeaderMismatchException duplicateHeader(String headerName) {
        return new HeaderMismatchException(headerName, "duplicate");
    }

}