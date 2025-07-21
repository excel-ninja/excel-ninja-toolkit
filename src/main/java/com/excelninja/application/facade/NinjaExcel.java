package com.excelninja.application.facade;

import com.excelninja.domain.exception.DocumentConversionException;
import com.excelninja.domain.exception.HeaderMismatchException;
import com.excelninja.domain.model.*;
import com.excelninja.infrastructure.converter.DefaultConverter;
import com.excelninja.infrastructure.io.PoiWorkbookReader;
import com.excelninja.infrastructure.io.PoiWorkbookWriter;
import com.excelninja.infrastructure.metadata.EntityMetadata;
import com.excelninja.infrastructure.metadata.FieldMapping;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

public final class NinjaExcel {
    private static final Logger logger = Logger.getLogger(NinjaExcel.class.getName());
    private static final PoiWorkbookReader WORKBOOK_READER = new PoiWorkbookReader();
    private static final PoiWorkbookWriter WORKBOOK_WRITER = new PoiWorkbookWriter();
    private static final DefaultConverter CONVERTER = new DefaultConverter();

    private NinjaExcel() {}

    public static <T> List<T> read(
            String filePath,
            Class<T> clazz
    ) {
        if (filePath == null || filePath.trim().isEmpty()) {
            throw new DocumentConversionException("File path cannot be null or empty");
        }
        return read(new File(filePath), clazz);
    }

    public static <T> List<T> read(
            File file,
            Class<T> clazz
    ) {
        validateReadInputs(file, clazz);

        long startTime = System.currentTimeMillis();
        String fileName = file.getName();
        long fileSize = file.length();

        logger.fine(String.format("[NINJA-EXCEL] Reading Excel file: %s (%.2f KB) [Cache size: %d]",
                fileName, fileSize / 1024.0, EntityMetadata.getCacheSize()));

        try {
            ExcelWorkbook workbook = WORKBOOK_READER.read(file);
            String firstSheetName = workbook.getSheetNames().iterator().next();
            ExcelSheet sheet = workbook.getSheet(firstSheetName);
            List<T> result = convertSheetToEntities(sheet, clazz);

            long duration = System.currentTimeMillis() - startTime;
            double recordsPerSecond = calculateRecordsPerSecond(result.size(), duration);

            logger.fine(String.format("[NINJA-EXCEL] Successfully read %d records from %s in %d ms (%.2f records/sec) [Cache size: %d]",
                    result.size(), fileName, duration, recordsPerSecond, EntityMetadata.getCacheSize()));

            return result;
        } catch (IOException e) {
            long duration = System.currentTimeMillis() - startTime;
            logger.severe(String.format("[NINJA-EXCEL] Failed to read Excel file: %s after %d ms", fileName, duration));
            throw new DocumentConversionException("Failed to read Excel file: " + file.getName() + ". Please check if the file exists and is not corrupted.", e);
        }
    }

    public static <T> List<T> readSheet(
            String filePath,
            String sheetName,
            Class<T> clazz
    ) {
        return readSheet(new File(filePath), sheetName, clazz);
    }

    public static <T> List<T> readSheet(
            File file,
            String sheetName,
            Class<T> clazz
    ) {
        validateReadInputs(file, clazz);

        long startTime = System.currentTimeMillis();
        String fileName = file.getName();

        logger.fine(String.format("[NINJA-EXCEL] Reading sheet '%s' from Excel file: %s [Cache size: %d]",
                sheetName, fileName, EntityMetadata.getCacheSize()));

        try {
            ExcelWorkbook workbook = WORKBOOK_READER.read(file);
            ExcelSheet sheet = workbook.getSheet(sheetName);

            if (sheet == null) {
                throw new DocumentConversionException("Sheet not found: " + sheetName);
            }

            List<T> result = convertSheetToEntities(sheet, clazz);

            long duration = System.currentTimeMillis() - startTime;
            double recordsPerSecond = calculateRecordsPerSecond(result.size(), duration);

            logger.fine(String.format("[NINJA-EXCEL] Successfully read %d records from sheet '%s' in %s in %d ms (%.2f records/sec) [Cache size: %d]",
                    result.size(), sheetName, fileName, duration, recordsPerSecond, EntityMetadata.getCacheSize()));

            return result;
        } catch (IOException e) {
            long duration = System.currentTimeMillis() - startTime;
            logger.severe(String.format("[NINJA-EXCEL] Failed to read Excel file: %s after %d ms", fileName, duration));
            throw new DocumentConversionException("Failed to read Excel file: " + file.getName(), e);
        }
    }

    public static <T> Map<String, List<T>> readAllSheets(
            String filePath,
            Class<T> clazz
    ) {
        return readAllSheets(new File(filePath), clazz);
    }

    public static <T> Map<String, List<T>> readAllSheets(
            File file,
            Class<T> clazz
    ) {
        validateReadInputs(file, clazz);

        long startTime = System.currentTimeMillis();
        String fileName = file.getName();

        logger.fine(String.format("[NINJA-EXCEL] Reading all sheets from Excel file: %s [Cache size: %d]",
                fileName, EntityMetadata.getCacheSize()));

        try {
            ExcelWorkbook workbook = WORKBOOK_READER.read(file);
            Map<String, List<T>> result = new LinkedHashMap<>();

            for (String sheetName : workbook.getSheetNames()) {
                ExcelSheet sheet = workbook.getSheet(sheetName);
                result.put(sheetName, convertSheetToEntities(sheet, clazz));
            }

            long duration = System.currentTimeMillis() - startTime;
            int totalRecords = result.values().stream().mapToInt(List::size).sum();

            logger.fine(String.format("[NINJA-EXCEL] Successfully read %d sheets with %d total records from %s in %d ms [Cache size: %d]",
                    result.size(), totalRecords, fileName, duration, EntityMetadata.getCacheSize()));

            return result;
        } catch (IOException e) {
            long duration = System.currentTimeMillis() - startTime;
            logger.severe(String.format("[NINJA-EXCEL] Failed to read Excel file: %s after %d ms", fileName, duration));
            throw new DocumentConversionException("Failed to read Excel file: " + file.getName(), e);
        }
    }

    public static <T> Map<String, List<T>> readSheets(
            String filePath,
            Class<T> clazz,
            List<String> sheetNames
    ) {
        return readSheets(new File(filePath), clazz, sheetNames);
    }

    public static <T> Map<String, List<T>> readSheets(
            File file,
            Class<T> clazz,
            List<String> sheetNames
    ) {
        validateReadInputs(file, clazz);

        long startTime = System.currentTimeMillis();
        String fileName = file.getName();

        logger.fine(String.format("[NINJA-EXCEL] Reading specified sheets %s from Excel file: %s [Cache size: %d]",
                sheetNames, fileName, EntityMetadata.getCacheSize()));

        try {
            ExcelWorkbook workbook = WORKBOOK_READER.read(file);
            Map<String, List<T>> result = new LinkedHashMap<>();

            for (String sheetName : sheetNames) {
                ExcelSheet sheet = workbook.getSheet(sheetName);
                if (sheet != null) {
                    result.put(sheetName, convertSheetToEntities(sheet, clazz));
                }
            }

            long duration = System.currentTimeMillis() - startTime;
            int totalRecords = result.values().stream().mapToInt(List::size).sum();

            logger.fine(String.format("[NINJA-EXCEL] Successfully read %d sheets with %d total records from %s in %d ms [Cache size: %d]",
                    result.size(), totalRecords, fileName, duration, EntityMetadata.getCacheSize()));

            return result;
        } catch (IOException e) {
            long duration = System.currentTimeMillis() - startTime;
            logger.severe(String.format("[NINJA-EXCEL] Failed to read Excel file: %s after %d ms", fileName, duration));
            throw new DocumentConversionException("Failed to read Excel file: " + file.getName(), e);
        }
    }

    public static List<String> getSheetNames(String filePath) {
        return getSheetNames(new File(filePath));
    }

    public static List<String> getSheetNames(File file) {
        validateReadInputs(file, Object.class);

        try {
            ExcelWorkbook workbook = WORKBOOK_READER.read(file);
            return new ArrayList<>(workbook.getSheetNames());
        } catch (IOException e) {
            throw new DocumentConversionException("Failed to read Excel file: " + file.getName(), e);
        }
    }

    public static void write(
            ExcelWorkbook workbook,
            String fileName
    ) {
        write(workbook, new File(fileName));
    }

    public static void write(
            ExcelWorkbook workbook,
            File file
    ) {
        if (workbook == null) {
            throw new DocumentConversionException("ExcelWorkbook cannot be null");
        }
        if (file == null) {
            throw new DocumentConversionException("File cannot be null");
        }

        long startTime = System.currentTimeMillis();
        String fileName = file.getName();
        int totalRecords = workbook.getSheetNames().stream()
                .mapToInt(sheetName -> workbook.getSheet(sheetName).getRows().size())
                .sum();

        logger.fine(String.format("[NINJA-EXCEL] Writing Excel workbook with %d sheets and %d total records to file: %s [Cache size: %d]",
                workbook.getSheetNames().size(), totalRecords, fileName, EntityMetadata.getCacheSize()));

        try {
            WORKBOOK_WRITER.write(workbook, file);

            long fileSize = file.length();
            long duration = System.currentTimeMillis() - startTime;
            double recordsPerSecond = calculateRecordsPerSecond(totalRecords, duration);

            logger.fine(String.format("[NINJA-EXCEL] Successfully wrote workbook with %d sheets and %d records to %s (%.2f KB) in %d ms (%.2f records/sec) [Cache size: %d]",
                    workbook.getSheetNames().size(), totalRecords, fileName, fileSize / 1024.0, duration, recordsPerSecond, EntityMetadata.getCacheSize()));

        } catch (IOException e) {
            long duration = System.currentTimeMillis() - startTime;
            logger.severe(String.format("[NINJA-EXCEL] Failed to write Excel workbook: %s after %d ms", fileName, duration));
            throw new DocumentConversionException("Failed to write Excel workbook: " + fileName + ". Please check file permissions and available disk space.", e);
        }
    }

    public static void write(
            ExcelWorkbook workbook,
            OutputStream outputStream
    ) {
        if (workbook == null) {
            throw new DocumentConversionException("ExcelWorkbook cannot be null");
        }
        if (outputStream == null) {
            throw new DocumentConversionException("OutputStream cannot be null");
        }

        long startTime = System.currentTimeMillis();
        int totalRecords = workbook.getSheetNames().stream()
                .mapToInt(sheetName -> workbook.getSheet(sheetName).getRows().size())
                .sum();

        logger.fine(String.format("[NINJA-EXCEL] Writing Excel workbook with %d sheets and %d total records to output stream [Cache size: %d]",
                workbook.getSheetNames().size(), totalRecords, EntityMetadata.getCacheSize()));

        try {
            WORKBOOK_WRITER.write(workbook, outputStream);

            long duration = System.currentTimeMillis() - startTime;
            double recordsPerSecond = calculateRecordsPerSecond(totalRecords, duration);

            logger.fine(String.format("[NINJA-EXCEL] Successfully wrote workbook with %d sheets and %d records to output stream in %d ms (%.2f records/sec) [Cache size: %d]",
                    workbook.getSheetNames().size(), totalRecords, duration, recordsPerSecond, EntityMetadata.getCacheSize()));

        } catch (IOException e) {
            long duration = System.currentTimeMillis() - startTime;
            logger.severe(String.format("[NINJA-EXCEL] Failed to write Excel workbook to output stream after %d ms", duration));
            throw new DocumentConversionException("Failed to write Excel workbook to output stream", e);
        }
    }

    private static <T> List<T> convertSheetToEntities(
            ExcelSheet sheet,
            Class<T> entityType
    ) {
        EntityMetadata<T> metadata = EntityMetadata.of(entityType);
        List<Integer> fieldToColumnMapping = createFieldToColumnMapping(sheet.getHeaders(), metadata);
        return convertRowsToEntities(sheet.getRows(), entityType, metadata, fieldToColumnMapping);
    }

    private static List<Integer> createFieldToColumnMapping(
            Headers headers,
            EntityMetadata<?> metadata
    ) {
        List<Integer> mapping = new ArrayList<>();
        for (FieldMapping fieldMapping : metadata.getReadFieldMappings()) {
            String headerName = fieldMapping.getHeaderName();
            if (!headers.containsHeader(headerName)) {
                throw HeaderMismatchException.headerNotFound(headerName);
            }
            mapping.add(headers.getPositionOf(headerName));
        }
        return mapping;
    }

    private static <T> List<T> convertRowsToEntities(
            DocumentRows rows,
            Class<T> entityType,
            EntityMetadata<T> metadata,
            List<Integer> fieldToColumnMapping
    ) {
        List<T> entities = new ArrayList<>();
        List<FieldMapping> fieldMappings = metadata.getReadFieldMappings();

        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            DocumentRow row = rows.getRow(rowIndex);

            try {
                T entity = metadata.createInstance();

                for (int fieldIndex = 0; fieldIndex < fieldMappings.size(); fieldIndex++) {
                    FieldMapping fieldMapping = fieldMappings.get(fieldIndex);
                    int columnIndex = fieldToColumnMapping.get(fieldIndex);

                    Object cellValue = row.getValue(columnIndex);
                    fieldMapping.setValue(entity, cellValue, CONVERTER);
                }

                entities.add(entity);

            } catch (Exception e) {
                throw new DocumentConversionException("Failed to create entity of type " + entityType.getName() + " for row " + (rowIndex + 1), e);
            }
        }

        return entities;
    }

    private static <T> void validateReadInputs(
            File file,
            Class<T> clazz
    ) {
        if (file == null) {
            throw new DocumentConversionException("File parameter cannot be null");
        }

        if (clazz == null) {
            throw new DocumentConversionException("Target class parameter cannot be null");
        }

        if (!file.exists()) {
            throw new DocumentConversionException("Excel file does not exist: " + file.getAbsolutePath());
        }

        if (!file.canRead()) {
            throw new DocumentConversionException("Cannot read Excel file (permission denied): " + file.getAbsolutePath());
        }

        if (file.length() == 0) {
            throw new DocumentConversionException("Excel file is empty: " + file.getAbsolutePath());
        }
    }

    private static double calculateRecordsPerSecond(
            int recordCount,
            long duration
    ) {
        return duration > 0 ? (recordCount * 1000.0 / duration) : 0;
    }

    static {
        logger.fine("[NINJA-EXCEL] Ninja Excel activated with unified API - simple reads, flexible writes!");
    }
}