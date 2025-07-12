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
import java.util.Iterator;
import java.util.List;

public class PoiExcelReader implements ExcelReader {

    @Override
    public ExcelDocument read(File excelFile) throws IOException {
        try (FileInputStream fileInputStream = new FileInputStream(excelFile)) {
            return read(fileInputStream);
        }
    }

    @Override
    public ExcelDocument read(InputStream inputStream) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {
            org.apache.poi.ss.usermodel.Sheet firstSheet = workbook.getSheetAt(0);
            String sheetName = firstSheet.getSheetName();
            Iterator<Row> rowIterator = firstSheet.iterator();

            if (!rowIterator.hasNext()) {
                throw new IllegalArgumentException("Excel sheet is empty");
            }

            Row headerRow = rowIterator.next();
            List<String> headerTitles = new ArrayList<String>();
            for (Cell cell : headerRow) {
                headerTitles.add(cell.getStringCellValue());
            }

            List<List<Object>> dataRows = new ArrayList<List<Object>>();
            while (rowIterator.hasNext()) {
                Row currentRow = rowIterator.next();
                List<Object> rowValues = new ArrayList<Object>();
                for (int columnIndex = 0; columnIndex < headerTitles.size(); columnIndex++) {
                    Cell cell = currentRow.getCell(columnIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    Object cellValue = extractCellValue(cell);
                    rowValues.add(cellValue);
                }
                dataRows.add(rowValues);
            }

            return ExcelDocument.builder()
                    .sheet(sheetName)
                    .headers(headerTitles)
                    .rows(dataRows)
                    .build();
        }
    }

    private Object extractCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return null;
            default:
                return cell.toString();
        }
    }
}
