package com.excelninja.infrastructure.io;

import com.excelninja.domain.exception.InvalidDocumentStructureException;
import com.excelninja.domain.model.ExcelSheet;
import com.excelninja.domain.model.ExcelWorkbook;
import com.excelninja.domain.port.WorkbookReader;
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

public class PoiWorkbookReader implements WorkbookReader {

    @Override
    public ExcelWorkbook read(File excelFile) throws IOException {
        try (FileInputStream fileInputStream = new FileInputStream(excelFile)) {
            return read(fileInputStream);
        }
    }

    @Override
    public ExcelWorkbook read(InputStream inputStream) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {
            ExcelWorkbook.WorkbookBuilder builder = ExcelWorkbook.builder();

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                org.apache.poi.ss.usermodel.Sheet poiSheet = workbook.getSheetAt(i);
                ExcelSheet sheet = readSheetFromPOI(poiSheet);
                builder.sheet(poiSheet.getSheetName(), sheet);
            }

            return builder.build();
        }
    }

    private ExcelSheet readSheetFromPOI(org.apache.poi.ss.usermodel.Sheet poiSheet) {
        String sheetName = poiSheet.getSheetName();
        Iterator<Row> rowIterator = poiSheet.iterator();

        if (!rowIterator.hasNext()) {
            throw new InvalidDocumentStructureException("Excel sheet is empty: " + sheetName);
        }

        Row headerRow = rowIterator.next();
        List<String> headerTitles = new ArrayList<>();
        for (Cell cell : headerRow) {
            String headerValue = cell.getStringCellValue();
            if (headerValue == null || headerValue.trim().isEmpty()) {
                throw new InvalidDocumentStructureException("Header cannot be empty at column " + cell.getColumnIndex() + " in sheet " + sheetName);
            }
            headerTitles.add(headerValue.trim());
        }

        List<List<Object>> dataRows = new ArrayList<>();
        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            List<Object> rowValues = new ArrayList<>();

            for (int columnIndex = 0; columnIndex < headerTitles.size(); columnIndex++) {
                Cell cell = currentRow.getCell(columnIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                Object cellValue = extractCellValue(cell);
                rowValues.add(cellValue);
            }

            dataRows.add(rowValues);
        }

        return ExcelSheet.builder()
                .name(sheetName)
                .headers(headerTitles)
                .rows(dataRows)
                .build();
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