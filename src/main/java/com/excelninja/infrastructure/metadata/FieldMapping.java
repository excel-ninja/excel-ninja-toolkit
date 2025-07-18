package com.excelninja.infrastructure.metadata;

import com.excelninja.application.port.ConverterPort;
import com.excelninja.domain.exception.DocumentConversionException;
import com.excelninja.infrastructure.util.ReflectionUtils;

import java.lang.reflect.Field;
import java.util.Objects;

public final class FieldMapping {

    public enum Type {
        READ, WRITE
    }

    private final Field field;
    private final String headerName;
    private final Class<?> targetType;
    private final String defaultValue;
    private final int order;
    private final Type type;

    public FieldMapping(
            Field field,
            String headerName,
            Class<?> targetType,
            String defaultValue,
            int order,
            Type type
    ) {
        this.field = Objects.requireNonNull(field, "Field cannot be null");
        this.headerName = Objects.requireNonNull(headerName, "Header name cannot be null");
        this.targetType = Objects.requireNonNull(targetType, "Target type cannot be null");
        this.defaultValue = defaultValue != null ? defaultValue : "";
        this.order = order;
        this.type = Objects.requireNonNull(type, "Type cannot be null");
    }

    public Object getValue(Object entity) {
        try {
            return ReflectionUtils.getFieldValue(entity, field);
        } catch (Exception e) {
            throw new DocumentConversionException("Failed to get value from field '" + field.getName() + "': " + e.getMessage(), e);
        }
    }

    public void setValue(
            Object entity,
            Object value,
            ConverterPort converter
    ) {
        try {
            Object convertedValue = convertValue(value, converter);
            ReflectionUtils.setFieldValue(entity, field, convertedValue);
        } catch (Exception e) {
            throw new DocumentConversionException("Failed to set value to field '" + field.getName() + "': " + e.getMessage(), e);
        }
    }

    private Object convertValue(
            Object value,
            ConverterPort converter
    ) {
        try {
            if (value != null) {
                return converter.convert(value, targetType);
            } else if (!defaultValue.isEmpty()) {
                return converter.convert(defaultValue, targetType);
            }
            return null;
        } catch (Exception e) {
            throw new DocumentConversionException(field.getName(), value, "Type conversion failed: " + e.getMessage());
        }
    }

    public Field getField() {
        return field;
    }

    public String getHeaderName() {
        return headerName;
    }

    public Class<?> getTargetType() {
        return targetType;
    }

    public String getDefaultValue() {
        return defaultValue;
    }

    public int getOrder() {
        return order;
    }

    public Type getType() {
        return type;
    }

    public String getFieldName() {
        return field.getName();
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        FieldMapping that = (FieldMapping) o;
        return order == that.order &&
                Objects.equals(field, that.field) &&
                Objects.equals(headerName, that.headerName) &&
                Objects.equals(targetType, that.targetType) &&
                Objects.equals(defaultValue, that.defaultValue) &&
                type == that.type;
    }

    @Override
    public int hashCode() {
        return Objects.hash(field, headerName, targetType, defaultValue, order, type);
    }

    @Override
    public String toString() {
        return String.format("FieldMapping{field=%s, header='%s', type=%s, order=%d}", field.getName(), headerName, targetType.getSimpleName(), order);
    }
}