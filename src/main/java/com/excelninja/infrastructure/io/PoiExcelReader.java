package com.excelninja.infrastructure.io;

import com.excelninja.domain.model.ExcelDocument;
import com.excelninja.domain.port.ExcelReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class PoiExcelReader implements ExcelReader {

    @Override
    public ExcelDocument read(File file) throws IOException {
        try (var is = new FileInputStream(file)) {
            return read(is);
        }
    }

    @Override
    public ExcelDocument read(InputStream inputStream) throws IOException {
        try (var workbook = new XSSFWorkbook(inputStream)) {
            var sheet = workbook.getSheetAt(0);
            var sheetName = sheet.getSheetName();
            var rowIterator = sheet.iterator();

            if (!rowIterator.hasNext()) {
                throw new IllegalArgumentException("Excel sheet is empty");
            }
            var headerRow = rowIterator.next();
            var headers = new ArrayList<String>();
            for (var cell : headerRow) {
                headers.add(cell.getStringCellValue());
            }

            var rows = new ArrayList<List<Object>>();
            while (rowIterator.hasNext()) {
                var row = rowIterator.next();
                var rowValues = new ArrayList<Object>();
                for (int i = 0; i < headers.size(); i++) {
                    var cell = row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    rowValues.add(getCellValue(cell));
                }
                rows.add(rowValues);
            }

            return ExcelDocument.builder()
                    .sheet(sheetName)
                    .headers(headers)
                    .rows(rows)
                    .build();
        }
    }

    private Object getCellValue(Cell cell) {
        if (cell == null) return null;
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> DateUtil.isCellDateFormatted(cell) ? cell.getDateCellValue() : cell.getNumericCellValue();
            case BOOLEAN -> cell.getBooleanCellValue();
            case FORMULA -> cell.getCellFormula();
            case BLANK -> null;
            default -> cell.toString();
        };
    }
}
