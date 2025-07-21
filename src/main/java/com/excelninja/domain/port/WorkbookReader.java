package com.excelninja.domain.port;

import com.excelninja.domain.model.ExcelWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

public interface WorkbookReader {
    ExcelWorkbook read(File excelFile) throws IOException;

    ExcelWorkbook read(InputStream inputStream) throws IOException;
}
