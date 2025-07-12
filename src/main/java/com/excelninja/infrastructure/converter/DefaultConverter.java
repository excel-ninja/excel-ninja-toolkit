package com.excelninja.infrastructure.converter;

import com.excelninja.application.port.ConverterPort;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

public class DefaultConverter implements ConverterPort {

    @Override
    public Object convert(Object rawValue, Class<?> targetType) {
        if (rawValue == null) {
            return null;
        }

        if (rawValue instanceof Number) {
            Number num = (Number) rawValue;
            if (targetType == int.class || targetType == Integer.class) {
                return num.intValue();
            }
            if (targetType == long.class || targetType == Long.class) {
                return num.longValue();
            }
            if (targetType == double.class || targetType == Double.class) {
                return num.doubleValue();
            }
        }

        if (rawValue instanceof Boolean && (targetType == boolean.class || targetType == Boolean.class)) {
            return rawValue;
        }

        if (targetType == LocalDate.class && rawValue instanceof String) {
            return LocalDate.parse((String) rawValue, DateTimeFormatter.ISO_DATE);
        }

        if (targetType == String.class) {
            return rawValue.toString();
        }

        if (targetType.isInstance(rawValue)) {
            return rawValue;
        }

        throw new IllegalArgumentException("Cannot convert value [" + rawValue + "] to " + targetType.getName());
    }
}
