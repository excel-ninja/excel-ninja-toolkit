package com.excelninja.infrastructure.util;

import com.excelninja.domain.exception.DocumentConversionException;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;

public class ReflectionUtils {

    private static final ConcurrentMap<String, Method> METHOD_CACHE = new ConcurrentHashMap<>();

    public static void setFieldValue(
            Object instance,
            Field field,
            Object value
    ) {
        if (instance == null) {
            throw new DocumentConversionException("Instance cannot be null");
        }

        if (field == null) {
            throw new DocumentConversionException("Field cannot be null");
        }

        if (trySetterMethod(instance, field, value)) {
            return;
        }

        setFieldValueDirectly(instance, field, value);
    }

    public static Object getFieldValue(
            Object instance,
            Field field
    ) {
        if (instance == null) {
            throw new DocumentConversionException("Instance cannot be null");
        }

        if (field == null) {
            throw new DocumentConversionException("Field cannot be null");
        }

        Object value = tryGetterMethod(instance, field);
        if (value != null || hasGetterMethod(instance, field)) {
            return value;
        }

        return getFieldValueDirectly(instance, field);
    }

    private static boolean trySetterMethod(
            Object instance,
            Field field,
            Object value
    ) {
        try {
            String setterName = "set" + capitalize(field.getName());
            String methodKey = instance.getClass().getName() + "." + setterName;

            Method setter = METHOD_CACHE.computeIfAbsent(methodKey, k -> {
                try {
                    return instance.getClass().getMethod(setterName, field.getType());
                } catch (NoSuchMethodException e) {
                    return null;
                }
            });

            if (setter != null) {
                if (!setter.isAccessible()) {
                    try {
                        setter.setAccessible(true);
                    } catch (SecurityException e) {
                        return false;
                    }
                }

                setter.invoke(instance, value);
                return true;
            }
        } catch (Exception e) {
            return false;
        }
        return false;
    }

    private static Object tryGetterMethod(
            Object instance,
            Field field
    ) {
        try {
            String getterName = getGetterName(field);
            String methodKey = instance.getClass().getName() + "." + getterName;

            Method getter = METHOD_CACHE.computeIfAbsent(methodKey, k -> {
                try {
                    return instance.getClass().getMethod(getterName);
                } catch (NoSuchMethodException e) {
                    return null;
                }
            });

            if (getter != null) {
                if (!getter.isAccessible()) {
                    try {
                        getter.setAccessible(true);
                    } catch (SecurityException e) {
                        return null;
                    }
                }

                return getter.invoke(instance);
            }
        } catch (Exception e) {
            return null;
        }
        return null;
    }

    private static boolean hasGetterMethod(
            Object instance,
            Field field
    ) {
        try {
            String getterName = getGetterName(field);
            instance.getClass().getMethod(getterName);
            return true;
        } catch (NoSuchMethodException e) {
            return false;
        }
    }

    private static void setFieldValueDirectly(
            Object instance,
            Field field,
            Object value
    ) {
        try {
            if (Modifier.isFinal(field.getModifiers())) {
                throw new DocumentConversionException("Cannot set final field: " + field.getName() + ". Consider adding a constructor parameter or setter method.");
            }

            if (!field.isAccessible()) {
                try {
                    field.setAccessible(true);
                } catch (SecurityException e) {
                    throw new DocumentConversionException("Cannot access field: " + field.getName() + ". Field is not accessible and cannot be made accessible. " + "Consider adding a public setter method: set" + capitalize(field.getName()) + "()", e);
                }
            }

            if (value != null && !isAssignable(field.getType(), value.getClass())) {
                throw new DocumentConversionException("Type mismatch for field: " + field.getName() + ". Expected: " + field.getType().getSimpleName() + ", but got: " + value.getClass().getSimpleName());
            }

            field.set(instance, value);

        } catch (IllegalAccessException e) {
            throw new DocumentConversionException("Failed to set field value: " + field.getName() + ". Consider adding a setter method: set" + capitalize(field.getName()) + "()", e);
        } catch (IllegalArgumentException e) {
            throw new DocumentConversionException("Invalid argument for field: " + field.getName() + ". Value: " + value + ", Expected type: " + field.getType().getSimpleName(), e);
        } catch (Exception e) {
            throw new DocumentConversionException("Unexpected error setting field: " + field.getName(), e);
        }
    }

    private static Object getFieldValueDirectly(
            Object instance,
            Field field
    ) {
        try {
            if (!field.isAccessible()) {
                try {
                    field.setAccessible(true);
                } catch (SecurityException e) {
                    throw new DocumentConversionException("Cannot access field: " + field.getName() + ". Field is not accessible and cannot be made accessible. " + "Consider adding a public getter method: " + getGetterName(field) + "()", e);
                }
            }

            return field.get(instance);

        } catch (IllegalAccessException e) {
            throw new DocumentConversionException("Failed to get field value: " + field.getName() + ". Consider adding a getter method: " + getGetterName(field) + "()", e);
        } catch (Exception e) {
            throw new DocumentConversionException("Unexpected error getting field: " + field.getName(), e);
        }
    }

    private static String getGetterName(Field field) {
        String fieldName = field.getName();

        if (field.getType() == boolean.class || field.getType() == Boolean.class) {
            if (fieldName.startsWith("is") && fieldName.length() > 2 &&
                    Character.isUpperCase(fieldName.charAt(2))) {
                return fieldName;
            }
            return "is" + capitalize(fieldName);
        }

        return "get" + capitalize(fieldName);
    }

    private static String capitalize(String str) {
        if (str == null || str.isEmpty()) {
            return str;
        }
        return str.substring(0, 1).toUpperCase() + str.substring(1);
    }

    private static boolean isAssignable(
            Class<?> targetType,
            Class<?> sourceType
    ) {
        if (targetType.isAssignableFrom(sourceType)) {
            return true;
        }

        if (targetType.isPrimitive()) {
            return isWrapperType(sourceType, targetType);
        }

        if (sourceType.isPrimitive()) {
            return isWrapperType(targetType, sourceType);
        }

        return false;
    }

    private static boolean isWrapperType(
            Class<?> wrapperType,
            Class<?> primitiveType
    ) {
        if (primitiveType == int.class && wrapperType == Integer.class) return true;
        if (primitiveType == long.class && wrapperType == Long.class) return true;
        if (primitiveType == double.class && wrapperType == Double.class) return true;
        if (primitiveType == float.class && wrapperType == Float.class) return true;
        if (primitiveType == boolean.class && wrapperType == Boolean.class) return true;
        if (primitiveType == byte.class && wrapperType == Byte.class) return true;
        if (primitiveType == char.class && wrapperType == Character.class) return true;
        return primitiveType == short.class && wrapperType == Short.class;
    }

    public static void clearMethodCache() {
        METHOD_CACHE.clear();
    }

    public static int getCachedMethodCount() {
        return METHOD_CACHE.size();
    }
}