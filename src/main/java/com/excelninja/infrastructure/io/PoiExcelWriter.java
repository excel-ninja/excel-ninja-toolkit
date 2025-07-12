package com.excelninja.infrastructure.io;

import com.excelninja.application.port.ConverterPort;
import com.excelninja.domain.model.ExcelDocument;
import com.excelninja.domain.port.ExcelWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.OutputStream;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.IntStream;

public class PoiExcelWriter implements ExcelWriter {

    @Override
    public void write(
            ExcelDocument doc,
            OutputStream out,
            ConverterPort converter
    ) {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            XSSFSheet sheet = workbook.createSheet(doc.getSheetName());

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

            XSSFRow headerRow = sheet.createRow(0);

            headerRow.setHeightInPoints(20);
            AtomicInteger columnIndex = new AtomicInteger();
            doc.getHeaders().forEach(title -> {
                XSSFCell cell = headerRow.createCell(columnIndex.getAndIncrement(), CellType.STRING);
                cell.setCellValue(title);
                cell.setCellStyle(headerStyle);
            });

            AtomicInteger rowIndex = new AtomicInteger(1);
            for (List<Object> rowValues : doc.getRows()) {
                XSSFRow row = sheet.createRow(rowIndex.getAndIncrement());
                AtomicInteger cellIndex = new AtomicInteger();
                rowValues.forEach(raw -> {
                    XSSFCell cell = row.createCell(cellIndex.getAndIncrement());
                    Object val = converter.convert(raw, raw != null ? raw.getClass() : String.class);
                    if (val == null) {
                        cell.setCellValue("");
                    } else if (val instanceof Number) {
                        Number numberValue = (Number) val;
                        cell.setCellValue(numberValue.doubleValue());
                    } else if (val instanceof Boolean) {
                        Boolean booleanValue = (Boolean) val;
                        cell.setCellValue(booleanValue);
                    } else {
                        cell.setCellValue(val.toString());
                    }
                    cell.setCellStyle(dataStyle);
                });
            }

            IntStream.range(0, doc.getHeaders().size()).forEach(column -> {
                if (doc.getColumnWidths().containsKey(column)) {
                    sheet.setColumnWidth(column, doc.getColumnWidths().get(column));
                } else {
                    sheet.autoSizeColumn(column);
                }
            });

            doc.getRowHeights().forEach((rowNumber, height) -> {
                XSSFRow row = sheet.getRow(rowNumber);
                if (row != null) {
                    row.setHeight(height);
                }
            });

            workbook.write(out);
        } catch (Exception e) {
            throw new RuntimeException("Failed to write Excel file", e);
        }
    }
}
