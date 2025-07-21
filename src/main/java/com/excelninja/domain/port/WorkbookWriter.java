package com.excelninja.domain.port;

import com.excelninja.domain.model.ExcelWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

public interface WorkbookWriter {
    void write(
            ExcelWorkbook workbook,
            File file
    ) throws IOException;

    void write(
            ExcelWorkbook workbook,
            OutputStream outputStream
    ) throws IOException;
}
