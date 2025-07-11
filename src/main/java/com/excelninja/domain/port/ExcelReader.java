package com.excelninja.domain.port;

import com.excelninja.domain.model.ExcelDocument;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

public interface ExcelReader {
    ExcelDocument read(File file) throws IOException;

    ExcelDocument read(InputStream inputStream) throws IOException;
}
