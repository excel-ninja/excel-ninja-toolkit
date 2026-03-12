package com.excelninja.infrastructure.io;

import com.excelninja.domain.exception.InvalidDocumentStructureException;
import com.excelninja.domain.model.ExcelSheet;
import com.excelninja.domain.model.ExcelWorkbook;
import com.excelninja.domain.model.WorkbookMetadata;
import com.excelninja.domain.port.WorkbookReader;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Apache POI-based Excel workbook reader.
 *
 * <p><b>Thread Safety:</b> This class is stateless and thread-safe.
 * Multiple threads can safely use the same instance concurrently.
 */
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
            ExcelWorkbook.WorkbookBuilder builder = ExcelWorkbook.builder()
                    .metadata(readWorkbookMetadata(workbook.getProperties().getCoreProperties()));

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                org.apache.poi.ss.usermodel.Sheet poiSheet = workbook.getSheetAt(i);
                ExcelSheet sheet = readSheetFromPOI(poiSheet);
                builder.sheet(poiSheet.getSheetName(), sheet);
            }

            return builder.build();
        }
    }

    public ExcelSheet readFirstSheet(File excelFile) throws IOException {
        try (FileInputStream fileInputStream = new FileInputStream(excelFile);
             XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream)) {
            if (workbook.getNumberOfSheets() == 0) {
                throw new InvalidDocumentStructureException("No sheets found in workbook");
            }

            return readSheetFromPOI(workbook.getSheetAt(0));
        }
    }

    public ExcelSheet readSheet(
            File excelFile,
            String sheetName
    ) throws IOException {
        try (FileInputStream fileInputStream = new FileInputStream(excelFile);
             XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream)) {
            org.apache.poi.ss.usermodel.Sheet poiSheet = workbook.getSheet(sheetName);
            if (poiSheet == null) {
                return null;
            }
            return readSheetFromPOI(poiSheet);
        }
    }

    public List<String> getSheetNames(File excelFile) throws IOException {
        try (FileInputStream fileInputStream = new FileInputStream(excelFile);
             XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream)) {
            List<String> sheetNames = new ArrayList<>(workbook.getNumberOfSheets());
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheetNames.add(workbook.getSheetName(i));
            }
            return sheetNames;
        }
    }

    public List<ExcelSheet> readSheets(
            File excelFile,
            List<String> requestedSheetNames
    ) throws IOException {
        try (FileInputStream fileInputStream = new FileInputStream(excelFile);
             XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream)) {
            List<ExcelSheet> sheets = new ArrayList<>();
            for (String sheetName : requestedSheetNames) {
                org.apache.poi.ss.usermodel.Sheet poiSheet = workbook.getSheet(sheetName);
                if (poiSheet != null) {
                    sheets.add(readSheetFromPOI(poiSheet));
                }
            }
            return sheets;
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
        int headerColumnCount = headerRow.getLastCellNum();
        if (headerColumnCount < 0) {
            throw new InvalidDocumentStructureException("No headers found in sheet: " + sheetName);
        }

        for (int columnIndex = 0; columnIndex < headerColumnCount; columnIndex++) {
            Cell cell = headerRow.getCell(columnIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            String headerValue = getCellValueAsString(cell);
            if (headerValue == null || headerValue.trim().isEmpty()) {
                throw new InvalidDocumentStructureException("Header cannot be empty at column " + columnIndex + " in sheet " + sheetName);
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

            if (hasMeaningfulValues(rowValues)) {
                dataRows.add(rowValues);
            }
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
                return extractFormulaValue(cell);
            case BLANK:
                return null;
            default:
                return cell.toString();
        }
    }

    private Object extractFormulaValue(Cell cell) {
        switch (cell.getCachedFormulaResultType()) {
            case STRING:
                String stringValue = cell.getStringCellValue();
                return stringValue == null || stringValue.isEmpty() ? null : stringValue;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                }
                return cell.getNumericCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case ERROR:
                return FormulaError.forInt(cell.getErrorCellValue()).getString();
            case BLANK:
                return null;
            default:
                return cell.toString();
        }
    }

    private boolean hasMeaningfulValues(List<Object> rowValues) {
        return rowValues.stream().anyMatch(this::hasMeaningfulValue);
    }

    private boolean hasMeaningfulValue(Object value) {
        if (value == null) {
            return false;
        }

        if (value instanceof String) {
            String stringValue = (String) value;
            return stringValue.isEmpty() || !stringValue.trim().isEmpty();
        }

        return true;
    }

    /**
     * Converts any cell value to a string representation.
     * This is useful for reading headers that may be numbers, formulas, or other types.
     *
     * @param cell the cell to convert
     * @return the string representation of the cell value, or null if the cell is blank
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return null;
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return formatValueAsString(extractCellValue(cell));
            case BOOLEAN:
                return formatValueAsString(cell.getBooleanCellValue());
            case FORMULA:
                return formatValueAsString(extractFormulaValue(cell));
            case BLANK:
                return null;
            default:
                return formatValueAsString(cell.toString());
        }
    }

    private String formatValueAsString(Object value) {
        if (value == null) {
            return null;
        }

        if (value instanceof Number) {
            double numericValue = ((Number) value).doubleValue();
            if (numericValue == Math.floor(numericValue) && !Double.isInfinite(numericValue) && !Double.isNaN(numericValue)) {
                return String.valueOf(((Number) value).longValue());
            }
            return String.valueOf(numericValue);
        }

        if (value instanceof java.util.Date) {
            LocalDateTime dateTime = ((java.util.Date) value).toInstant()
                    .atZone(ZoneId.systemDefault())
                    .toLocalDateTime();
            if (dateTime.toLocalTime().equals(LocalTime.MIDNIGHT)) {
                return DateTimeFormatter.ISO_LOCAL_DATE.format(dateTime.toLocalDate());
            }
            return DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss")
                    .format(dateTime.withNano(0));
        }

        return value.toString();
    }

    private WorkbookMetadata readWorkbookMetadata(POIXMLProperties.CoreProperties coreProperties) {
        if (coreProperties == null) {
            return new WorkbookMetadata();
        }

        LocalDateTime createdDate = coreProperties.getCreated() != null
                ? LocalDateTime.ofInstant(coreProperties.getCreated().toInstant(), ZoneId.systemDefault())
                : null;

        return new WorkbookMetadata(
                coreProperties.getCreator(),
                coreProperties.getTitle(),
                createdDate
        );
    }
}
