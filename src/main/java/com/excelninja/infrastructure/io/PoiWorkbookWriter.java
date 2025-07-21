package com.excelninja.infrastructure.io;

import com.excelninja.domain.model.DocumentRow;
import com.excelninja.domain.model.ExcelSheet;
import com.excelninja.domain.model.ExcelWorkbook;
import com.excelninja.domain.model.Header;
import com.excelninja.domain.port.WorkbookWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;

public class PoiWorkbookWriter implements WorkbookWriter {

    @Override
    public void write(
            ExcelWorkbook workbook,
            File file
    ) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(file)) {
            write(workbook, fos);
        }
    }

    @Override
    public void write(
            ExcelWorkbook workbook,
            OutputStream outputStream
    ) throws IOException {
        try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {

            XSSFCellStyle dateStyle = createDateStyle(poiWorkbook);
            XSSFCellStyle dateTimeStyle = createDateTimeStyle(poiWorkbook);
            XSSFCellStyle headerStyle = createHeaderStyle(poiWorkbook);
            XSSFCellStyle dataStyle = createDataStyle(poiWorkbook);

            for (String sheetName : workbook.getSheetNames()) {
                ExcelSheet excelSheet = workbook.getSheet(sheetName);
                XSSFSheet poiSheet = poiWorkbook.createSheet(sheetName);

                createHeaderRow(poiSheet, excelSheet, headerStyle);
                createDataRows(poiSheet, excelSheet, dataStyle, dateStyle, dateTimeStyle);
                adjustColumnWidths(poiSheet, excelSheet);
                adjustRowHeights(poiSheet, excelSheet);
            }

            poiWorkbook.write(outputStream);
        }
    }

    private void createHeaderRow(
            XSSFSheet sheet,
            ExcelSheet excelSheet,
            XSSFCellStyle headerStyle
    ) {
        XSSFRow headerRow = sheet.createRow(0);
        headerRow.setHeightInPoints(20);

        for (Header header : excelSheet.getHeaders().getHeaders()) {
            XSSFCell cell = headerRow.createCell(header.getPosition(), CellType.STRING);
            cell.setCellValue(header.getName());
            cell.setCellStyle(headerStyle);
        }
    }

    private void createDataRows(
            XSSFSheet sheet,
            ExcelSheet excelSheet,
            XSSFCellStyle dataStyle,
            XSSFCellStyle dateStyle,
            XSSFCellStyle dateTimeStyle
    ) {
        for (DocumentRow documentRow : excelSheet.getRows().getRows()) {
            XSSFRow row = sheet.createRow(documentRow.getRowNumber());

            for (int columnIndex = 0; columnIndex < documentRow.getColumnCount(); columnIndex++) {
                XSSFCell cell = row.createCell(columnIndex);
                Object rawValue = documentRow.getValue(columnIndex);
                setCellValue(cell, rawValue, dataStyle, dateStyle, dateTimeStyle);
            }
        }
    }

    private void setCellValue(
            XSSFCell cell,
            Object rawValue,
            XSSFCellStyle dataStyle,
            XSSFCellStyle dateStyle,
            XSSFCellStyle dateTimeStyle
    ) {
        if (rawValue == null) {
            cell.setCellValue("");
            cell.setCellStyle(dataStyle);
        } else if (rawValue instanceof Number) {
            if (rawValue instanceof BigDecimal) {
                cell.setCellValue(((BigDecimal) rawValue).doubleValue());
            } else {
                cell.setCellValue(((Number) rawValue).doubleValue());
            }
            cell.setCellStyle(dataStyle);
        } else if (rawValue instanceof Boolean) {
            cell.setCellValue((Boolean) rawValue);
            cell.setCellStyle(dataStyle);
        } else if (rawValue instanceof LocalDateTime) {
            LocalDateTime localDateTime = (LocalDateTime) rawValue;
            Date date = Date.from(localDateTime.atZone(ZoneId.systemDefault()).toInstant());
            cell.setCellValue(date);
            cell.setCellStyle(dateTimeStyle);
        } else if (rawValue instanceof LocalDate) {
            LocalDate localDate = (LocalDate) rawValue;
            Date date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
            cell.setCellValue(date);
            cell.setCellStyle(dateStyle);
        } else if (rawValue instanceof Date) {
            cell.setCellValue((Date) rawValue);
            cell.setCellStyle(dateTimeStyle);
        } else {
            cell.setCellValue(rawValue.toString());
            cell.setCellStyle(dataStyle);
        }
    }

    private void adjustColumnWidths(
            XSSFSheet sheet,
            ExcelSheet excelSheet
    ) {
        for (int columnIndex = 0; columnIndex < excelSheet.getHeaders().size(); columnIndex++) {
            if (excelSheet.getMetadata().getColumnWidths().containsKey(columnIndex)) {
                sheet.setColumnWidth(columnIndex, excelSheet.getMetadata().getColumnWidths().get(columnIndex));
            } else if (excelSheet.getMetadata().isAutoSizeColumns()) {
                sheet.autoSizeColumn(columnIndex);
            }
        }
    }

    private void adjustRowHeights(
            XSSFSheet sheet,
            ExcelSheet excelSheet
    ) {
        excelSheet.getMetadata().getRowHeights().forEach((rowNumber, height) -> {
            XSSFRow row = sheet.getRow(rowNumber);
            if (row != null) {
                row.setHeight(height);
            }
        });
    }

    private XSSFCellStyle createHeaderStyle(XSSFWorkbook workbook) {
        XSSFCellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        headerStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        headerStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        headerStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return headerStyle;
    }

    private XSSFCellStyle createDataStyle(XSSFWorkbook workbook) {
        XSSFCellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setBorderBottom(BorderStyle.THIN);
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setBorderRight(BorderStyle.THIN);
        dataStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        dataStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        dataStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        dataStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return dataStyle;
    }

    private XSSFCellStyle createDateStyle(XSSFWorkbook workbook) {
        XSSFCellStyle dateStyle = workbook.createCellStyle();
        dateStyle.setBorderTop(BorderStyle.THIN);
        dateStyle.setBorderBottom(BorderStyle.THIN);
        dateStyle.setBorderLeft(BorderStyle.THIN);
        dateStyle.setBorderRight(BorderStyle.THIN);
        dateStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        dateStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        dateStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        dateStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        dateStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        XSSFDataFormat dataFormat = workbook.createDataFormat();
        dateStyle.setDataFormat(dataFormat.getFormat("yyyy-mm-dd"));
        return dateStyle;
    }

    private XSSFCellStyle createDateTimeStyle(XSSFWorkbook workbook) {
        XSSFCellStyle dateTimeStyle = workbook.createCellStyle();
        dateTimeStyle.setBorderTop(BorderStyle.THIN);
        dateTimeStyle.setBorderBottom(BorderStyle.THIN);
        dateTimeStyle.setBorderLeft(BorderStyle.THIN);
        dateTimeStyle.setBorderRight(BorderStyle.THIN);
        dateTimeStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        dateTimeStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        dateTimeStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        dateTimeStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        dateTimeStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        XSSFDataFormat dataFormat = workbook.createDataFormat();
        dateTimeStyle.setDataFormat(dataFormat.getFormat("yyyy-mm-dd hh:mm:ss"));
        return dateTimeStyle;
    }
}