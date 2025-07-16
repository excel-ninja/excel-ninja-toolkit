package com.excelninja.domain.model;

import com.excelninja.domain.exception.InvalidDocumentStructureException;

import java.util.Objects;

public final class SheetName {

    private static final int MAX_LENGTH = 31;
    private static final String INVALID_CHARS = "/\\?*[]";

    private final String value;

    public SheetName(String value) {
        validateSheetName(value);
        this.value = value.trim();
    }

    private void validateSheetName(String value) {
        if (value == null || value.trim().isEmpty()) {
            throw InvalidDocumentStructureException.emptySheetName();
        }

        String trimmed = value.trim();

        if (trimmed.length() > MAX_LENGTH) {
            throw new InvalidDocumentStructureException(String.format("Sheet name too long: %d characters. Maximum allowed: %d", trimmed.length(), MAX_LENGTH));
        }

        for (char c : INVALID_CHARS.toCharArray()) {
            if (trimmed.indexOf(c) >= 0) {
                throw new InvalidDocumentStructureException(String.format("Sheet name contains invalid character '%s'. " + "Invalid characters: %s", c, INVALID_CHARS));
            }
        }

        if (trimmed.startsWith("'") || trimmed.endsWith("'")) {
            throw new InvalidDocumentStructureException("Sheet name cannot start or end with single quote");
        }

        // Check for reserved names
        if ("History".equalsIgnoreCase(trimmed)) {
            throw new InvalidDocumentStructureException("Sheet name 'History' is reserved and cannot be used");
        }
    }

    public String getValue() {
        return value;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        SheetName sheetName = (SheetName) o;
        return Objects.equals(value, sheetName.value);
    }

    @Override
    public int hashCode() {
        return Objects.hash(value);
    }

    @Override
    public String toString() {
        return value;
    }

    public static SheetName defaultName() {
        return new SheetName("Sheet1");
    }

    public boolean isDefault() {
        return "Sheet1".equals(value);
    }
}
