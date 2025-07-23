package com.excelninja.application.facade;

import com.excelninja.domain.exception.DocumentConversionException;
import com.excelninja.domain.exception.EntityMappingException;
import com.excelninja.domain.exception.HeaderMismatchException;
import com.excelninja.domain.model.*;
import com.excelninja.domain.port.WorkbookReader;
import com.excelninja.infrastructure.converter.DefaultConverter;
import com.excelninja.infrastructure.io.PoiWorkbookReader;
import com.excelninja.infrastructure.io.PoiWorkbookWriter;
import com.excelninja.infrastructure.io.StreamingWorkbookReader;
import com.excelninja.infrastructure.metadata.EntityMetadata;
import com.excelninja.infrastructure.metadata.FieldMapping;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;

public final class NinjaExcel {
    private static final Logger logger = Logger.getLogger(NinjaExcel.class.getName());
    private static final PoiWorkbookReader POI_WORKBOOK_READER = new PoiWorkbookReader();
    private static final StreamingWorkbookReader STREAMING_WORKBOOK_READER = new StreamingWorkbookReader();
    private static final PoiWorkbookWriter WORKBOOK_WRITER = new PoiWorkbookWriter();
    private static final DefaultConverter CONVERTER = new DefaultConverter();

    private static final long STREAMING_THRESHOLD_BYTES = 10 * 1024 * 1024; // 10MB
    private static final int DEFAULT_CHUNK_SIZE = 1000;

    private NinjaExcel() {}

    public static void setStreamingThreshold(long thresholdBytes) {
        logger.info(String.format("[NINJA-EXCEL] Streaming threshold updated to %.2f MB",
                thresholdBytes / (1024.0 * 1024.0)));
    }

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
        boolean useStreaming = fileSize > 1024 * 1024; // 1MB for simplicity

        logger.info(String.format("[NINJA-EXCEL] Reading Excel file: %s (%.2f MB) using %s reader",
                fileName, fileSize / (1024.0 * 1024.0),
                useStreaming ? "STREAMING" : "POI"));

        try {
            if (useStreaming) {
                return readWithStreaming(file, clazz);
            } else {
                return readWithPoi(file, clazz);
            }
        } catch (EntityMappingException | HeaderMismatchException e) {
            throw e;
        } catch (Exception e) {
            long duration = System.currentTimeMillis() - startTime;
            logger.log(Level.SEVERE, String.format("[NINJA-EXCEL] Failed to read Excel file: %s after %d ms", fileName, duration), e);
            throw new DocumentConversionException("Failed to read Excel file: " + file.getName() +
                    ". Please check if the file exists and is not corrupted.", e);
        }
        // =================================================================
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
        long fileSize = file.length();
        boolean useStreaming = shouldUseStreaming(fileSize);

        logger.info(String.format("[NINJA-EXCEL] Reading sheet '%s' from Excel file: %s (%.2f MB) using %s reader [Cache size: %d]",
                sheetName, fileName, fileSize / (1024.0 * 1024.0),
                useStreaming ? "STREAMING" : "POI",
                EntityMetadata.getCacheSize()));

        try {
            WorkbookReader reader = useStreaming ? STREAMING_WORKBOOK_READER : POI_WORKBOOK_READER;
            ExcelWorkbook workbook = reader.read(file);
            ExcelSheet sheet = workbook.getSheet(sheetName);

            if (sheet == null) {
                throw new DocumentConversionException("Sheet not found: " + sheetName);
            }

            List<T> result = convertSheetToEntities(sheet, clazz);

            long duration = System.currentTimeMillis() - startTime;
            double recordsPerSecond = calculateRecordsPerSecond(result.size(), duration);

            logger.info(String.format("[NINJA-EXCEL] Successfully read %d records from sheet '%s' in %s in %d ms (%.2f records/sec) using %s [Cache size: %d]",
                    result.size(), sheetName, fileName, duration, recordsPerSecond,
                    useStreaming ? "STREAMING" : "POI", EntityMetadata.getCacheSize()));

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
        long fileSize = file.length();
        boolean useStreaming = shouldUseStreaming(fileSize);

        logger.info(String.format("[NINJA-EXCEL] Reading all sheets from Excel file: %s (%.2f MB) using %s reader [Cache size: %d]",
                fileName, fileSize / (1024.0 * 1024.0),
                useStreaming ? "STREAMING" : "POI",
                EntityMetadata.getCacheSize()));

        try {
            WorkbookReader reader = useStreaming ? STREAMING_WORKBOOK_READER : POI_WORKBOOK_READER;
            ExcelWorkbook workbook = reader.read(file);
            Map<String, List<T>> result = new LinkedHashMap<>();

            for (String sheetName : workbook.getSheetNames()) {
                ExcelSheet sheet = workbook.getSheet(sheetName);
                result.put(sheetName, convertSheetToEntities(sheet, clazz));
            }

            long duration = System.currentTimeMillis() - startTime;
            int totalRecords = result.values().stream().mapToInt(List::size).sum();

            logger.info(String.format("[NINJA-EXCEL] Successfully read %d sheets with %d total records from %s in %d ms using %s [Cache size: %d]",
                    result.size(), totalRecords, fileName, duration,
                    useStreaming ? "STREAMING" : "POI", EntityMetadata.getCacheSize()));

            return result;
        } catch (IOException e) {
            long duration = System.currentTimeMillis() - startTime;
            logger.severe(String.format("[NINJA-EXCEL] Failed to read Excel file: %s after %d ms", fileName, duration));
            throw new DocumentConversionException("Failed to read Excel file: " + file.getName(), e);
        }
    }

    public static <T> Iterator<List<T>> readInChunks(
            String filePath,
            Class<T> clazz
    ) {
        return readInChunks(new File(filePath), clazz, DEFAULT_CHUNK_SIZE);
    }

    public static <T> Iterator<List<T>> readInChunks(
            File file,
            Class<T> clazz
    ) {
        return readInChunks(file, clazz, DEFAULT_CHUNK_SIZE);
    }

    public static <T> Iterator<List<T>> readInChunks(
            String filePath,
            Class<T> clazz,
            int chunkSize
    ) {
        return readInChunks(new File(filePath), clazz, chunkSize);
    }

    public static <T> Iterator<List<T>> readInChunks(
            File file,
            Class<T> clazz,
            int chunkSize
    ) {
        validateReadInputs(file, clazz);

        if (chunkSize <= 0) {
            throw new DocumentConversionException("Chunk size must be positive");
        }

        long fileSize = file.length();
        String fileName = file.getName();

        logger.info(String.format("[NINJA-EXCEL] Creating chunk iterator for Excel file: %s (%.2f MB) with chunk size: %d",
                fileName, fileSize / (1024.0 * 1024.0), chunkSize));

        try {
            return STREAMING_WORKBOOK_READER.readInChunks(file, clazz, chunkSize);
        } catch (IOException e) {
            throw new DocumentConversionException("Failed to create chunk iterator for file: " + fileName, e);
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
        long fileSize = file.length();
        boolean useStreaming = shouldUseStreaming(fileSize);

        logger.info(String.format("[NINJA-EXCEL] Reading specified sheets %s from Excel file: %s (%.2f MB) using %s reader [Cache size: %d]",
                sheetNames, fileName, fileSize / (1024.0 * 1024.0),
                useStreaming ? "STREAMING" : "POI",
                EntityMetadata.getCacheSize()));

        try {
            WorkbookReader reader = useStreaming ? STREAMING_WORKBOOK_READER : POI_WORKBOOK_READER;
            ExcelWorkbook workbook = reader.read(file);
            Map<String, List<T>> result = new LinkedHashMap<>();

            for (String sheetName : sheetNames) {
                ExcelSheet sheet = workbook.getSheet(sheetName);
                if (sheet != null) {
                    result.put(sheetName, convertSheetToEntities(sheet, clazz));
                }
            }

            long duration = System.currentTimeMillis() - startTime;
            int totalRecords = result.values().stream().mapToInt(List::size).sum();

            logger.info(String.format("[NINJA-EXCEL] Successfully read %d sheets with %d records from %s in %d ms using %s [Cache size: %d]",
                    result.size(), totalRecords, fileName, duration,
                    useStreaming ? "STREAMING" : "POI", EntityMetadata.getCacheSize()));

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
            long fileSize = file.length();
            boolean useStreaming = shouldUseStreaming(fileSize);
            WorkbookReader reader = useStreaming ? STREAMING_WORKBOOK_READER : POI_WORKBOOK_READER;

            ExcelWorkbook workbook = reader.read(file);
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

    private static boolean shouldUseStreaming(long fileSize) {
        boolean useStreaming = fileSize > STREAMING_THRESHOLD_BYTES;
        logger.fine(String.format(
                "[NINJA-EXCEL] File size: %.2f MB, threshold: %.2f MB, using %s reader",
                fileSize / (1024.0 * 1024.0),
                STREAMING_THRESHOLD_BYTES / (1024.0 * 1024.0),
                useStreaming ? "STREAMING" : "POI"
        ));
        return useStreaming;
    }

    private static <T> List<T> readWithPoi(
            File file,
            Class<T> clazz
    ) throws IOException {
        ExcelWorkbook workbook = POI_WORKBOOK_READER.read(file);
        String firstSheetName = workbook.getSheetNames().iterator().next();
        ExcelSheet sheet = workbook.getSheet(firstSheetName);
        return convertSheetToEntities(sheet, clazz);
    }

    private static <T> List<T> readWithStreaming(
            File file,
            Class<T> clazz
    ) throws IOException {
        ExcelWorkbook workbook = STREAMING_WORKBOOK_READER.read(file);
        String firstSheetName = workbook.getSheetNames().iterator().next();
        ExcelSheet sheet = workbook.getSheet(firstSheetName);
        return convertSheetToEntities(sheet, clazz);
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
}