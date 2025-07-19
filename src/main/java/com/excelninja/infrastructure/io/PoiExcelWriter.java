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

            XSSFCellStyle headerStyle = createHeaderStyle(workbook);
            XSSFCellStyle dataStyle = createDataStyle(workbook);

            createHeaderRow(sheet, doc, headerStyle);

            createDataRows(sheet, doc, converter, dataStyle);

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
            XSSFCellStyle dataStyle
    ) {
        for (DocumentRow documentRow : doc.getRows().getRows()) {
            XSSFRow row = sheet.createRow(documentRow.getRowNumber());

            for (int columnIndex = 0; columnIndex < documentRow.getColumnCount(); columnIndex++) {
                XSSFCell cell = row.createCell(columnIndex);
                Object rawValue = documentRow.getValue(columnIndex);
                setCellValue(cell, rawValue, dataStyle);
            }
        }
    }

    private void setCellValue(
            XSSFCell cell,
            Object rawValue,
            XSSFCellStyle dataStyle
    ) {
        if (rawValue == null) {
            cell.setCellValue("");
        } else if (rawValue instanceof Number) {
            if (rawValue instanceof BigDecimal) {
                cell.setCellValue(((BigDecimal) rawValue).doubleValue());
            } else {
                cell.setCellValue(((Number) rawValue).doubleValue());
            }
        } else if (rawValue instanceof Boolean) {
            cell.setCellValue((Boolean) rawValue);
        } else if (rawValue instanceof LocalDateTime) {
            LocalDateTime localDateTime = (LocalDateTime) rawValue;
            Date date = Date.from(localDateTime.atZone(ZoneId.systemDefault()).toInstant());
            cell.setCellValue(date);
        } else if (rawValue instanceof LocalDate) {
            LocalDate localDate = (LocalDate) rawValue;
            Date date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
            cell.setCellValue(date);
        } else if (rawValue instanceof Date) {
            cell.setCellValue((Date) rawValue);
        } else {
            cell.setCellValue(rawValue.toString());
        }

        cell.setCellStyle(dataStyle);
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
}