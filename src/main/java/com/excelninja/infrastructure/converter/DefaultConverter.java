package com.excelninja.infrastructure.converter;

import com.excelninja.application.port.ConverterPort;
import com.excelninja.domain.exception.DocumentConversionException;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;

public class DefaultConverter implements ConverterPort {

    @Override
    public Object convert(Object rawValue, Class<?> targetType) {
        if (rawValue == null) {
            return null;
        }

        try {
            if (rawValue instanceof Number) {
                return convertNumber((Number) rawValue, targetType);
            }

            if (rawValue instanceof Boolean && (targetType == boolean.class || targetType == Boolean.class)) {
                return rawValue;
            }

            if (targetType == LocalDate.class && rawValue instanceof String) {
                return convertToLocalDate((String) rawValue);
            }

            if (targetType == String.class) {
                return rawValue.toString();
            }

            if (targetType.isInstance(rawValue)) {
                return rawValue;
            }

            throw new DocumentConversionException(String.format("Cannot convert value '%s' of type %s to %s", rawValue, rawValue.getClass().getSimpleName(), targetType.getSimpleName()));

        } catch (DocumentConversionException e) {
            throw e;
        } catch (Exception e) {
            throw new DocumentConversionException(String.format("Unexpected error converting value '%s' to %s", rawValue, targetType.getSimpleName()), e);
        }
    }

    private Object convertNumber(
            Number num,
            Class<?> targetType
    ) {
        if (targetType == int.class || targetType == Integer.class) {
            if (num.doubleValue() > Integer.MAX_VALUE || num.doubleValue() < Integer.MIN_VALUE) {
                throw new DocumentConversionException(String.format("Number %s is out of Integer range (%d to %d)", num, Integer.MIN_VALUE, Integer.MAX_VALUE));
            }
            return num.intValue();
        }

        if (targetType == long.class || targetType == Long.class) {
            if (num.doubleValue() > Long.MAX_VALUE || num.doubleValue() < Long.MIN_VALUE) {
                throw new DocumentConversionException(String.format("Number %s is out of Long range (%d to %d)", num, Long.MIN_VALUE, Long.MAX_VALUE));
            }
            return num.longValue();
        }

        if (targetType == double.class || targetType == Double.class) {
            return num.doubleValue();
        }

        return num;
    }

    private LocalDate convertToLocalDate(String dateStr) {
        try {
            return LocalDate.parse(dateStr, DateTimeFormatter.ISO_DATE);
        } catch (DateTimeParseException e) {
            DateTimeFormatter[] formatters = {
                    DateTimeFormatter.ofPattern("yyyy-MM-dd"),
                    DateTimeFormatter.ofPattern("dd/MM/yyyy"),
                    DateTimeFormatter.ofPattern("MM/dd/yyyy"),
                    DateTimeFormatter.ofPattern("yyyy/MM/dd"),
                    DateTimeFormatter.ofPattern("dd-MM-yyyy"),
                    DateTimeFormatter.ofPattern("MM-dd-yyyy")
            };

            for (DateTimeFormatter formatter : formatters) {
                try {
                    return LocalDate.parse(dateStr, formatter);
                } catch (DateTimeParseException ignored) {
                }
            }

            throw new DocumentConversionException(String.format("Cannot parse date string '%s'. Supported formats: yyyy-MM-dd, dd/MM/yyyy, MM/dd/yyyy, etc.", dateStr));
        }
    }
}