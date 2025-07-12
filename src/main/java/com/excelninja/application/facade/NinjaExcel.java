package com.excelninja.application.facade;

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
        try (FileOutputStream out = new FileOutputStream(fileName)) {
            write(document, out);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> List<T> read(
            File file,
            Class<T> clazz
    ) {
        try {
            ExcelDocument document = READER.read(file);
            return document.toDTO(clazz, CONVERTER);
        } catch (IOException e) {
            throw new RuntimeException(e);
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
