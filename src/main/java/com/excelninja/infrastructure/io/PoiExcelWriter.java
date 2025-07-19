package com.excelninja.infrastructure.io;

import com.excelninja.application.port.ConverterPort;
import com.excelninja.domain.model.DocumentRow;
import com.excelninja.domain.model.ExcelDocument;
import com.excelninja.domain.model.Header;
import com.excelninja.domain.port.ExcelWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.OutputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;
import java.util.stream.IntStream;

public class PoiExcelWriter implements ExcelWriter {

    @Override
    public void write(
            ExcelDocument doc,
            OutputStream out,
            ConverterPort converter
    ) {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            XSSFSheet sheet = workbook.createSheet(doc.getSheetName().getValue());

            XSSFCellStyle dateStyle = createDateStyle(workbook);
            XSSFCellStyle dateTimeStyle = createDateTimeStyle(workbook);
            XSSFCellStyle headerStyle = createHeaderStyle(workbook);
            XSSFCellStyle dataStyle = createDataStyle(workbook);

            createHeaderRow(sheet, doc, headerStyle);

            createDataRows(sheet, doc, converter, dataStyle, dateStyle, dateTimeStyle);

            adjustColumnWidths(sheet, doc);
            adjustRowHeights(sheet, doc);

            workbook.write(out);
        } catch (Exception e) {
            throw new RuntimeException("Failed to write Excel file", e);
        }
    }

    private void createHeaderRow(
            XSSFSheet sheet,
            ExcelDocument doc,
            XSSFCellStyle headerStyle
    ) {
        XSSFRow headerRow = sheet.createRow(0);
        headerRow.setHeightInPoints(20);

        for (Header header : doc.getHeaders().getHeaders()) {
            XSSFCell cell = headerRow.createCell(header.getPosition(), CellType.STRING);
            cell.setCellValue(header.getName());
            cell.setCellStyle(headerStyle);
        }
    }

    private void createDataRows(
            XSSFSheet sheet,
            ExcelDocument doc,
            ConverterPort converter,
            XSSFCellStyle dataStyle,
            XSSFCellStyle dateStyle,
            XSSFCellStyle dateTimeStyle
    ) {
        for (DocumentRow documentRow : doc.getRows().getRows()) {
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
            ExcelDocument doc
    ) {
        IntStream.range(0, doc.getHeaders().size()).forEach(columnIndex -> {
            if (doc.getColumnWidths().containsKey(columnIndex)) {
                sheet.setColumnWidth(columnIndex, doc.getColumnWidths().get(columnIndex));
            } else {
                sheet.autoSizeColumn(columnIndex);
            }
        });
    }

    private void adjustRowHeights(
            XSSFSheet sheet,
            ExcelDocument doc
    ) {
        doc.getRowHeights().forEach((rowNumber, height) -> {
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