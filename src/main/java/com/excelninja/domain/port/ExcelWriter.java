package com.excelninja.domain.port;

import com.excelninja.application.port.ConverterPort;
import com.excelninja.domain.model.ExcelDocument;

import java.io.OutputStream;

public interface ExcelWriter {
    void write(
            ExcelDocument doc,
            OutputStream out,
            ConverterPort converter
    );
}
