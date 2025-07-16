package com.excelninja.application.facade;

import com.excelninja.domain.exception.DocumentConversionException;
import com.excelninja.domain.model.ExcelDocument;
import com.excelninja.domain.port.ExcelReader;
import com.excelninja.domain.port.ExcelWriter;
import com.excelninja.infrastructure.converter.DefaultConverter;
import com.excelninja.infrastructure.io.PoiExcelReader;
import com.excelninja.infrastructure.io.PoiExcelWriter;

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
            File file,
            Class<T> clazz
    ) {
        validateReadInputs(file, clazz);
        try {
            ExcelDocument document = READER.read(file);
            return document.convertToEntities(clazz, CONVERTER);
        } catch (IOException e) {
            throw new DocumentConversionException("Failed to read Excel file: " + file.getName() + ". Please check if the file exists and is not corrupted.", e);
        }
    }

    public static void write(
            ExcelDocument document,
            OutputStream out
    ) {
        validateWriteInputs(document, out);

        try {
            WRITER.write(document, out, CONVERTER);
        } catch (Exception e) {
            throw new DocumentConversionException("Failed to write Excel document to output stream", e);
        }
    }

    public static void write(
            ExcelDocument document,
            String fileName
    ) {
        validateWriteInputs(document, fileName);
        try (FileOutputStream out = new FileOutputStream(fileName)) {
            write(document, out);
        } catch (IOException e) {
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

    static {
        logger.info(
                "         ██████      / /\n" +
                        "              █████████ / /       \n" +
                        "            █  ───  ───  █/        \n" +
                        "             ██████████          \n" +
                        "                 ██████ /            \n" +
                        "─────██───────██─────\n" +
                        "🥷  E X C E L - N I N J A"
        );
    }
}
