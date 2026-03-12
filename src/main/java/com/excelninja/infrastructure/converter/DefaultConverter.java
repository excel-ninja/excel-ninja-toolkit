package com.excelninja.infrastructure.converter;

import com.excelninja.application.port.ConverterPort;
import com.excelninja.domain.exception.DocumentConversionException;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Date;

/**
 * Default type converter for Excel cell values.
 *
 * <p><b>Thread Safety:</b> This class is stateless and thread-safe.
 * Multiple threads can safely use the same instance concurrently.
 */
public class DefaultConverter implements ConverterPort {
    private static final DateTimeFormatter ISO_LOCAL_DATE_FORMATTER = DateTimeFormatter.ISO_LOCAL_DATE;
    private static final DateTimeFormatter ISO_LOCAL_DATE_TIME_SECONDS_FORMATTER =
            DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss");

    @Override
    public Object convert(Object rawValue, Class<?> targetType) {
        if (rawValue == null) {
            return null;
        }

        try {
            if (targetType.isInstance(rawValue)) {
                return rawValue;
            }

            if (rawValue instanceof Number) {
                return convertNumber((Number) rawValue, targetType);
            }

            if (rawValue instanceof Boolean && (targetType == boolean.class || targetType == Boolean.class)) {
                return rawValue;
            }

            if (targetType == LocalDate.class) {
                return convertToLocalDate(rawValue);
            }

            if (targetType == LocalDateTime.class) {
                return convertToLocalDateTime(rawValue);
            }

            if (targetType == Date.class) {
                return convertToDate(rawValue);
            }

            if (targetType == BigDecimal.class) {
                return convertToBigDecimal(rawValue);
            }

            if (targetType == String.class) {
                return convertToString(rawValue);
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

        if (targetType == float.class || targetType == Float.class) {
            return num.floatValue();
        }

        if (targetType == BigDecimal.class) {
            return convertNumberToBigDecimal(num);
        }

        return num;
    }

    private BigDecimal convertNumberToBigDecimal(Number num) {
        if (num instanceof BigDecimal) {
            return (BigDecimal) num;
        }

        if (num instanceof Double) {
            return new BigDecimal(num.toString());
        }

        if (num instanceof Float) {
            return new BigDecimal(num.toString());
        }

        if (num instanceof Long || num instanceof Integer ||
                num instanceof Short || num instanceof Byte) {
            return BigDecimal.valueOf(num.longValue());
        }

        return BigDecimal.valueOf(num.doubleValue());
    }

    private LocalDate convertToLocalDate(Object value) {
        if (value instanceof LocalDate) {
            return (LocalDate) value;
        }

        if (value instanceof Date) {
            return ((Date) value).toInstant()
                    .atZone(ZoneId.systemDefault())
                    .toLocalDate();
        }

        if (value instanceof String) {
            return parseLocalDateFromString((String) value);
        }

        if (value instanceof LocalDateTime) {
            return ((LocalDateTime) value).toLocalDate();
        }

        throw new DocumentConversionException(String.format("Cannot convert value '%s' of type %s to LocalDate", value, value.getClass().getSimpleName()));
    }

    private LocalDateTime convertToLocalDateTime(Object value) {
        if (value instanceof LocalDateTime) {
            return (LocalDateTime) value;
        }

        if (value instanceof Date) {
            return ((Date) value).toInstant()
                    .atZone(ZoneId.systemDefault())
                    .toLocalDateTime();
        }

        if (value instanceof String) {
            return parseLocalDateTimeFromString((String) value);
        }

        if (value instanceof LocalDate) {
            return ((LocalDate) value).atStartOfDay();
        }

        throw new DocumentConversionException(String.format("Cannot convert value '%s' of type %s to LocalDateTime", value, value.getClass().getSimpleName()));
    }

    private Date convertToDate(Object value) {
        if (value instanceof Date) {
            return (Date) value;
        }

        if (value instanceof LocalDate) {
            return Date.from(((LocalDate) value).atStartOfDay(ZoneId.systemDefault()).toInstant());
        }

        if (value instanceof LocalDateTime) {
            return Date.from(((LocalDateTime) value).atZone(ZoneId.systemDefault()).toInstant());
        }

        if (value instanceof String) {
            String stringValue = (String) value;
            try {
                LocalDateTime localDateTime = parseLocalDateTimeFromString(stringValue);
                return Date.from(localDateTime.atZone(ZoneId.systemDefault()).toInstant());
            } catch (DocumentConversionException ignored) {
            }

            try {
                LocalDate localDate = parseLocalDateFromString(stringValue);
                return Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
            } catch (DocumentConversionException ignored) {
            }
        }

        throw new DocumentConversionException(String.format("Cannot convert value '%s' of type %s to Date", value, value.getClass().getSimpleName()));
    }

    private String convertToString(Object value) {
        if (value instanceof Date) {
            return formatDateTime(((Date) value).toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime());
        }

        if (value instanceof LocalDate) {
            return ((LocalDate) value).format(ISO_LOCAL_DATE_FORMATTER);
        }

        if (value instanceof LocalDateTime) {
            return formatDateTime((LocalDateTime) value);
        }

        return value.toString();
    }

    private String formatDateTime(LocalDateTime localDateTime) {
        if (localDateTime.toLocalTime().equals(LocalTime.MIDNIGHT)) {
            return localDateTime.toLocalDate().format(ISO_LOCAL_DATE_FORMATTER);
        }

        return localDateTime.withNano(0).format(ISO_LOCAL_DATE_TIME_SECONDS_FORMATTER);
    }

    private BigDecimal convertToBigDecimal(Object value) {
        if (value instanceof BigDecimal) {
            return (BigDecimal) value;
        }

        if (value instanceof Number) {
            return convertNumberToBigDecimal((Number) value);
        }

        if (value instanceof String) {
            try {
                return new BigDecimal((String) value);
            } catch (NumberFormatException e) {
                throw new DocumentConversionException(String.format("Cannot parse '%s' as BigDecimal", value));
            }
        }

        throw new DocumentConversionException(String.format("Cannot convert value '%s' of type %s to BigDecimal", value, value.getClass().getSimpleName()));
    }

    private LocalDate parseLocalDateFromString(String dateStr) {
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

    private LocalDateTime parseLocalDateTimeFromString(String dateTimeStr) {
        try {
            return LocalDateTime.parse(dateTimeStr, DateTimeFormatter.ISO_LOCAL_DATE_TIME);
        } catch (DateTimeParseException e) {
            DateTimeFormatter[] formatters = {
                    DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"),
                    DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm"),
                    DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss"),
                    DateTimeFormatter.ofPattern("MM/dd/yyyy HH:mm:ss"),
                    DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss"),
                    DateTimeFormatter.ofPattern("dd-MM-yyyy HH:mm:ss"),
                    DateTimeFormatter.ofPattern("MM-dd-yyyy HH:mm:ss"),
                    DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss")
            };

            for (DateTimeFormatter formatter : formatters) {
                try {
                    return LocalDateTime.parse(dateTimeStr, formatter);
                } catch (DateTimeParseException ignored) {
                }
            }

            throw new DocumentConversionException(String.format("Cannot parse datetime string '%s'. Supported formats: yyyy-MM-dd HH:mm:ss, yyyy-MM-dd'T'HH:mm:ss, etc.", dateTimeStr));
        }
    }
}
