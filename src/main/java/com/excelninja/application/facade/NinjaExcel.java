package com.excelninja.application.facade;

import com.excelninja.domain.exception.DocumentConversionException;
import com.excelninja.domain.model.ExcelDocument;
import com.excelninja.domain.port.ExcelReader;
import com.excelninja.domain.port.ExcelWriter;
import com.excelninja.infrastructure.converter.DefaultConverter;
import com.excelninja.infrastructure.io.PoiExcelReader;
import com.excelninja.infrastructure.io.PoiExcelWriter;
import com.excelninja.infrastructure.metadata.EntityMetadata;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.logging.Logger;

public final class NinjaExcel {
    private static final Logger logger = Logger.getLogger(NinjaExcel.class.getName());
    private static final ExcelWriter WRITER = new PoiExcelWriter();
    private static final ExcelReader READER = new PoiExcelReader();
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

        logger.fine(String.format(
                "[NINJA-EXCEL] Reading Excel file: %s (%.2f KB) [Cache size: %d]",
                fileName,
                fileSize / 1024.0,
                EntityMetadata.getCacheSize()
        ));

        try {
            ExcelDocument document = READER.read(file);
            List<T> result = document.convertToEntities(clazz, CONVERTER);

            long duration = System.currentTimeMillis() - startTime;
            double recordsPerSecond = calculateRecordsPerSecond(result.size(), duration);

            logger.fine(String.format(
                    "[NINJA-EXCEL] Successfully read %d records from %s in %d ms (%.2f records/sec) [Cache size: %d]",
                    result.size(),
                    fileName,
                    duration,
                    recordsPerSecond,
                    EntityMetadata.getCacheSize()
            ));

            return result;
        } catch (IOException e) {
            long duration = System.currentTimeMillis() - startTime;
            logger.severe(String.format("[NINJA-EXCEL] Failed to read Excel file: %s after %d ms", fileName, duration));

            throw new DocumentConversionException("Failed to read Excel file: " + file.getName() + ". Please check if the file exists and is not corrupted.", e);
        }
    }

    public static void write(
            ExcelDocument document,
            OutputStream out
    ) {
        validateWriteInputs(document, out);

        long startTime = System.currentTimeMillis();
        int recordCount = document.getRowCount();

        logger.fine(String.format(
                "[NINJA-EXCEL] Writing Excel document with %d records to output stream [Cache size: %d]",
                recordCount,
                EntityMetadata.getCacheSize()
        ));

        try {
            WRITER.write(document, out, CONVERTER);

            long duration = System.currentTimeMillis() - startTime;
            double recordsPerSecond = calculateRecordsPerSecond(recordCount, duration);

            logger.fine(String.format(
                    "[NINJA-EXCEL] Successfully wrote %d records to output stream in %d ms (%.2f records/sec) [Cache size: %d]",
                    recordCount,
                    duration,
                    recordsPerSecond,
                    EntityMetadata.getCacheSize()
            ));

        } catch (Exception e) {
            long duration = System.currentTimeMillis() - startTime;
            logger.severe(String.format("[NINJA-EXCEL] Failed to write Excel document to output stream after %d ms", duration));

            throw new DocumentConversionException("Failed to write Excel document to output stream", e);
        }
    }

    public static void write(
            ExcelDocument document,
            String fileName
    ) {
        validateWriteInputs(document, fileName);

        long startTime = System.currentTimeMillis();
        int recordCount = document.getRowCount();

        logger.fine(String.format(
                "[NINJA-EXCEL] Writing Excel document with %d records to file: %s [Cache size: %d]",
                recordCount,
                fileName,
                EntityMetadata.getCacheSize()
        ));

        try (FileOutputStream out = new FileOutputStream(fileName)) {
            WRITER.write(document, out, CONVERTER);

            File writtenFile = new File(fileName);
            long fileSize = writtenFile.length();
            long duration = System.currentTimeMillis() - startTime;
            double recordsPerSecond = calculateRecordsPerSecond(recordCount, duration);

            logger.fine(String.format(
                    "[NINJA-EXCEL] Successfully wrote %d records to %s (%.2f KB) in %d ms (%.2f records/sec) [Cache size: %d]",
                    recordCount,
                    fileName,
                    fileSize / 1024.0,
                    duration,
                    recordsPerSecond,
                    EntityMetadata.getCacheSize()
            ));

        } catch (IOException e) {
            long duration = System.currentTimeMillis() - startTime;
            logger.severe(String.format("[NINJA-EXCEL] Failed to write Excel file: %s after %d ms", fileName, duration));

            throw new DocumentConversionException("Failed to write Excel file: " + fileName + ". Please check file permissions and available disk space.", e);
        }
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

    private static void validateWriteInputs(
            ExcelDocument document,
            Object output
    ) {
        if (document == null) {
            throw new DocumentConversionException("ExcelDocument cannot be null");
        }

        if (output == null) {
            throw new DocumentConversionException("Output parameter cannot be null");
        }

        if (output instanceof String) {
            String fileName = (String) output;
            if (fileName.trim().isEmpty()) {
                throw new DocumentConversionException("File name cannot be empty");
            }

            File file = new File(fileName);
            File parentDir = file.getParentFile();
            if (parentDir != null && !parentDir.exists()) {
                throw new DocumentConversionException("Directory does not exist: " + parentDir.getAbsolutePath());
            }
        }
    }

    private static double calculateRecordsPerSecond(
            int recordCount,
            long duration
    ) {
        return duration > 0 ? (recordCount * 1000.0 / duration) : 0;
    }

    static {
        logger.fine("[NINJA-EXCEL] Ninja Excel activated with metadata caching!");
    }
}