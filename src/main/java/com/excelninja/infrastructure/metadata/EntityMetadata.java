package com.excelninja.infrastructure.metadata;

import com.excelninja.domain.annotation.ExcelReadColumn;
import com.excelninja.domain.annotation.ExcelWriteColumn;
import com.excelninja.domain.exception.EntityMappingException;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.Collectors;

public class EntityMetadata<T> {
    private static final Map<Class<?>, EntityMetadata<?>> METADATA_CACHE = new ConcurrentHashMap<>();

    private final Class<T> entityType;
    private final Constructor<T> defaultConstructor;
    private final List<FieldMapping> readFieldMappings;
    private final List<FieldMapping> writeFieldMappings;
    private final Map<String, FieldMapping> headerToFieldMap;

    private EntityMetadata(Class<T> entityType) {
        this.entityType = entityType;
        this.readFieldMappings = extractReadFieldMappings(entityType);
        this.writeFieldMappings = extractWriteFieldMappings(entityType);
        this.headerToFieldMap = createHeaderToFieldMap(readFieldMappings);

        if (readFieldMappings.isEmpty() && writeFieldMappings.isEmpty()) {
            throw EntityMappingException.noAnnotatedFields(entityType);
        }

        if (!readFieldMappings.isEmpty()) {
            this.defaultConstructor = extractDefaultConstructor(entityType);
        } else {
            this.defaultConstructor = null;
        }
    }

    @SuppressWarnings("unchecked")
    public static <T> EntityMetadata<T> of(Class<T> entityType) {
        return (EntityMetadata<T>) METADATA_CACHE.computeIfAbsent(entityType,
                key -> new EntityMetadata<>(entityType));
    }

    public T createInstance() {
        if (defaultConstructor == null) {
            throw new EntityMappingException(entityType,
                    "Cannot create instance - no default constructor available. This entity has no read fields or no default constructor.");
        }

        try {
            return defaultConstructor.newInstance();
        } catch (Exception e) {
            throw new EntityMappingException(entityType,
                    "Failed to create instance using default constructor: " + e.getMessage());
        }
    }

    public List<FieldMapping> getReadFieldMappings() {
        return readFieldMappings;
    }

    public List<FieldMapping> getWriteFieldMappings() {
        return writeFieldMappings;
    }

    public Optional<FieldMapping> getFieldMappingByHeader(String headerName) {
        return Optional.ofNullable(headerToFieldMap.get(headerName));
    }

    public List<String> getReadHeaders() {
        return readFieldMappings.stream()
                .map(FieldMapping::getHeaderName)
                .collect(Collectors.toList());
    }

    public List<String> getWriteHeaders() {
        return writeFieldMappings.stream()
                .sorted(Comparator.comparing(FieldMapping::getOrder))
                .map(FieldMapping::getHeaderName)
                .collect(Collectors.toList());
    }

    public static int getCacheSize() {
        return METADATA_CACHE.size();
    }

    public static void clearCache() {
        METADATA_CACHE.clear();
    }

    public static void evictCache(Class<?> entityType) {
        METADATA_CACHE.remove(entityType);
    }

    private Constructor<T> extractDefaultConstructor(Class<T> entityType) {
        try {
            Constructor<T> constructor = entityType.getDeclaredConstructor();
            constructor.setAccessible(true);
            return constructor;
        } catch (NoSuchMethodException e) {
            throw new EntityMappingException(entityType, "No default constructor found. Entity classes must have a public or accessible no-args constructor for read operations.");
        } catch (SecurityException e) {
            throw new EntityMappingException(entityType, "Cannot access default constructor due to security restrictions.");
        }
    }

    private List<FieldMapping> extractReadFieldMappings(Class<T> entityType) {
        List<FieldMapping> mappings = new ArrayList<>();

        for (Field field : getAllFields(entityType)) {
            ExcelReadColumn annotation = field.getAnnotation(ExcelReadColumn.class);
            if (annotation != null) {
                field.setAccessible(true);
                mappings.add(new FieldMapping(
                        field,
                        annotation.headerName(),
                        annotation.type() == Void.class ? field.getType() : annotation.type(),
                        annotation.defaultValue(),
                        0,
                        FieldMapping.Type.READ
                ));
            }
        }

        return Collections.unmodifiableList(mappings);
    }

    private List<FieldMapping> extractWriteFieldMappings(Class<T> entityType) {
        List<FieldMapping> mappings = new ArrayList<>();

        for (Field field : getAllFields(entityType)) {
            ExcelWriteColumn annotation = field.getAnnotation(ExcelWriteColumn.class);
            if (annotation != null) {
                field.setAccessible(true);
                mappings.add(new FieldMapping(
                        field,
                        annotation.headerName(),
                        field.getType(),
                        "",
                        annotation.order(),
                        FieldMapping.Type.WRITE
                ));
            }
        }

        if (!mappings.isEmpty()) {
            validateWriteMappings(mappings, entityType);
            mappings.sort(Comparator.comparing(FieldMapping::getOrder));
        }

        return Collections.unmodifiableList(mappings);
    }

    private List<Field> getAllFields(Class<?> clazz) {
        List<Field> fields = new ArrayList<>();
        Class<?> currentClass = clazz;

        while (currentClass != null && currentClass != Object.class) {
            fields.addAll(Arrays.asList(currentClass.getDeclaredFields()));
            currentClass = currentClass.getSuperclass();
        }

        return fields;
    }

    private void validateWriteMappings(
            List<FieldMapping> mappings,
            Class<T> entityType
    ) {
        Set<String> usedHeaders = new HashSet<>();
        Set<Integer> usedOrders = new HashSet<>();

        for (FieldMapping mapping : mappings) {
            String headerName = mapping.getHeaderName();
            if (headerName == null || headerName.trim().isEmpty()) {
                throw EntityMappingException.invalidAnnotationConfiguration(entityType, "Empty header name in field: " + mapping.getField().getName());
            }

            if (!usedHeaders.add(headerName)) {
                throw EntityMappingException.invalidAnnotationConfiguration(entityType, "Duplicate header name: " + headerName);
            }

            int order = mapping.getOrder();
            if (order >= 0 && order != Integer.MAX_VALUE && !usedOrders.add(order)) {
                throw EntityMappingException.invalidAnnotationConfiguration(entityType, "Duplicate order value: " + order);
            }
        }
    }

    private Map<String, FieldMapping> createHeaderToFieldMap(List<FieldMapping> mappings) {
        Map<String, FieldMapping> map = new HashMap<>();
        for (FieldMapping mapping : mappings) {
            map.put(mapping.getHeaderName(), mapping);
        }
        return Collections.unmodifiableMap(map);
    }

    @Override
    public String toString() {
        return String.format(
                "EntityMetadata{entityType=%s, readFields=%d, writeFields=%d}",
                entityType.getSimpleName(),
                readFieldMappings.size(),
                writeFieldMappings.size()
        );
    }
}