package com.excelninja.domain.exception;

public abstract class ExcelDomainException extends RuntimeException {

    protected ExcelDomainException(String message) {
        super(message);
    }

    protected ExcelDomainException(
            String message,
            Throwable cause
    ) {
        super(message, cause);
    }
}