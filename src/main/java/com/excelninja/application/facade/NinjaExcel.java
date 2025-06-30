package com.excelninja.application.facade;

import com.excelninja.domain.model.ExcelDocument;
import com.excelninja.infrastructure.converter.DefaultConverter;
import com.excelninja.infrastructure.persistence.PoiExcelWriter;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public final class NinjaExcel {
    private static final PoiExcelWriter WRITER = new PoiExcelWriter();
    private static final DefaultConverter CONVERTER = new DefaultConverter();

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
}
