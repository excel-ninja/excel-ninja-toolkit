package com.excelninja.domain.exception;

public class EntityMappingException extends ExcelDomainException {

    private final Class<?> entityType;

    public EntityMappingException(
            Class<?> entityType,
            String message
    ) {
        super(String.format("Entity mapping failed for %s: %s", entityType.getSimpleName(), message));
        this.entityType = entityType;
    }

    public static EntityMappingException noAnnotatedFields(Class<?> entityType) {
        return new EntityMappingException(entityType, "No @ExcelReadColumn or @ExcelWriteColumn annotations found");
    }

    public static EntityMappingException invalidAnnotationConfiguration(
            Class<?> entityType,
            String details
    ) {
        return new EntityMappingException(entityType, "Invalid annotation configuration: " + details);
    }

    public static EntityMappingException emptyEntityList() {
        return new EntityMappingException(Object.class, "Entity list cannot be null or empty");
    }

}