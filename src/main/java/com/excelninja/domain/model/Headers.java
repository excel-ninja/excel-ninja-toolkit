package com.excelninja.domain.model;

import com.excelninja.domain.exception.HeaderMismatchException;
import com.excelninja.domain.exception.InvalidDocumentStructureException;

import java.util.*;
import java.util.stream.Collectors;

public final class Headers {

    private final List<Header> headers;
    private final Map<String, Integer> nameToPositionMap;

    public Headers(List<Header> headers) {
        validateHeaders(headers);
        this.headers = Collections.unmodifiableList(new ArrayList<>(headers));
        this.nameToPositionMap = createNameToPositionMap(headers);
    }

    private void validateHeaders(List<Header> headers) {
        if (headers == null || headers.isEmpty()) {
            throw InvalidDocumentStructureException.emptyHeaders();
        }

        Set<String> uniqueNames = new HashSet<>();
        Set<Integer> uniquePositions = new HashSet<>();

        for (Header header : headers) {
            if (header == null) {
                throw new InvalidDocumentStructureException("Header cannot be null");
            }

            if (!uniqueNames.add(header.getName())) {
                throw HeaderMismatchException.duplicateHeader(header.getName());
            }

            if (!uniquePositions.add(header.getPosition())) {
                throw new InvalidDocumentStructureException(
                        "Duplicate header position: " + header.getPosition());
            }
        }
    }

    private Map<String, Integer> createNameToPositionMap(List<Header> headers) {
        Map<String, Integer> map = new HashMap<>();
        for (Header header : headers) {
            map.put(header.getName(), header.getPosition());
        }
        return Collections.unmodifiableMap(map);
    }

    public List<Header> getHeaders() {
        return headers;
    }

    public int size() {
        return headers.size();
    }

    public boolean isEmpty() {
        return headers.isEmpty();
    }

    public Header getHeader(int position) {
        return headers.stream()
                .filter(header -> header.getPosition() == position)
                .findFirst()
                .orElseThrow(() -> new InvalidDocumentStructureException(
                        "No header found at position: " + position));
    }

    public Optional<Header> findHeader(String name) {
        return headers.stream()
                .filter(header -> header.getName().equals(name))
                .findFirst();
    }

    public boolean containsHeader(String name) {
        return nameToPositionMap.containsKey(name);
    }

    public int getPositionOf(String headerName) {
        Integer position = nameToPositionMap.get(headerName);
        if (position == null) {
            throw HeaderMismatchException.headerNotFound(headerName);
        }
        return position;
    }

    public List<String> getHeaderNames() {
        return headers.stream()
                .sorted(Comparator.comparing(Header::getPosition))
                .map(Header::getName)
                .collect(Collectors.toList());
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Headers headers1 = (Headers) o;
        return Objects.equals(headers, headers1.headers);
    }

    @Override
    public int hashCode() {
        return Objects.hash(headers);
    }

    @Override
    public String toString() {
        return "Headers{" + getHeaderNames() + "}";
    }

    public static Headers of(String... headerNames) {
        List<Header> headerList = new ArrayList<>();
        for (int i = 0; i < headerNames.length; i++) {
            headerList.add(new Header(headerNames[i], i));
        }
        return new Headers(headerList);
    }

    public static Headers of(List<String> headerNames) {
        List<Header> headerList = new ArrayList<>();
        for (int i = 0; i < headerNames.size(); i++) {
            headerList.add(new Header(headerNames.get(i), i));
        }
        return new Headers(headerList);
    }

    public Headers withAdditionalHeader(String name) {
        List<Header> newHeaders = new ArrayList<>(headers);
        newHeaders.add(new Header(name, headers.size()));
        return new Headers(newHeaders);
    }
}