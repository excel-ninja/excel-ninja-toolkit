package com.excelninja.application.exception;

import com.excelninja.domain.exception.ExcelDomainException;

public class ApplicationException extends RuntimeException {

    private final ExcelDomainException domainException;

    public ApplicationException(
            ExcelDomainException domainException,
            Throwable technicalCause
    ) {
        super(domainException.getMessage(), technicalCause);
        this.domainException = domainException;
    }

    public ExcelDomainException getDomainException() {
        return domainException;
    }

    public Throwable getTechnicalCause() {
        return super.getCause();
    }
}