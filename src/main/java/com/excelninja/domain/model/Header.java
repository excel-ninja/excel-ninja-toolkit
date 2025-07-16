package com.excelninja.domain.model;

import com.excelninja.domain.exception.InvalidDocumentStructureException;

import java.util.Objects;

public final class Header {

    private static final int MAX_HEADER_LENGTH = 255;

    private final String name;
    private final int position;

    public Header(
            String name,
            int position
    ) {
        validateHeader(name, position);
        this.name = name.trim();
        this.position = position;
    }

    private void validateHeader(
            String name,
            int position
    ) {
        if (name == null || name.trim().isEmpty()) {
            throw new InvalidDocumentStructureException("Header name cannot be empty");
        }

        String trimmed = name.trim();
        if (trimmed.length() > MAX_HEADER_LENGTH) {
            throw new InvalidDocumentStructureException(
                    String.format("Header name too long: %d characters. Maximum allowed: %d", trimmed.length(), MAX_HEADER_LENGTH));
        }

        if (position < 0) {
            throw new InvalidDocumentStructureException("Header position cannot be negative: " + position);
        }
    }

    public String getName() {
        return name;
    }

    public int getPosition() {
        return position;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        Header header = (Header) o;
        return position == header.position && Objects.equals(name, header.name);
    }

    @Override
    public int hashCode() {
        return Objects.hash(name, position);
    }

    @Override
    public String toString() {
        return String.format("Header{name='%s', position=%d}", name, position);
    }

    public static Header of(
            String name,
            int position
    ) {
        return new Header(name, position);
    }
}
