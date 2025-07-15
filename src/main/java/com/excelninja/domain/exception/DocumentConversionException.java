package com.excelninja.domain.exception;

public class DocumentConversionException extends ExcelDomainException {

    private final String fieldName;
    private final Object problematicValue;

    public DocumentConversionException(String message) {
        super(message);
        this.fieldName = null;
        this.problematicValue = null;
    }

    public DocumentConversionException(
            String message,
            Throwable cause
    ) {
        super(message, cause);
        this.fieldName = null;
        this.problematicValue = null;
    }

    public DocumentConversionException(
            String fieldName,
            Object value,
            String reason
    ) {
        super(String.format("Failed to convert field '%s' with value '%s': %s", fieldName, value, reason));
        this.fieldName = fieldName;
        this.problematicValue = value;
    }

}