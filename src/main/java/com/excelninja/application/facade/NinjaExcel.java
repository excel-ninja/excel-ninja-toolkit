package com.excelninja.application.facade;

import com.excelninja.domain.model.ExcelDocument;
import com.excelninja.infrastructure.converter.DefaultConverter;
import com.excelninja.infrastructure.persistence.PoiExcelReader;
import com.excelninja.infrastructure.persistence.PoiExcelWriter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public final class NinjaExcel<T> {
    private static final PoiExcelWriter WRITER = new PoiExcelWriter();
    private static final PoiExcelReader READER = new PoiExcelReader();
    private static final DefaultConverter CONVERTER = new DefaultConverter();
    private File file;
    private String sheetName;
    private OutputStream outputStream;
    private List<T> data;

    private NinjaExcel() {}

    public static void write(
            ExcelDocument document,
            OutputStream out
    ) {
        WRITER.write(document, out, CONVERTER);
    }

    public static void write(
            ExcelDocument document,
            String fileName
    ) {
        try (var out = new FileOutputStream(fileName)) {
            WRITER.write(document, out, CONVERTER);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> List<T> read(
            File file,
            Class<T> clazz
    ) {
        try {
            var document = READER.read(file);
            return document.toDTO(clazz, CONVERTER);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
