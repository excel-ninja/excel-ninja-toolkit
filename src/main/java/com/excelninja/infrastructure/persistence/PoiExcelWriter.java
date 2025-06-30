package com.excelninja.infrastructure.persistence;

import com.excelninja.application.port.ConverterPort;
import com.excelninja.domain.model.ExcelDocument;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.OutputStream;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.IntStream;

public class PoiExcelWriter {
    public void write(
            ExcelDocument doc,
            OutputStream out,
            ConverterPort converter
    ) {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet(doc.getSheetName());

            var headerStyle = workbook.createCellStyle();
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

            var dataStyle = workbook.createCellStyle();
            dataStyle.setBorderTop(BorderStyle.THIN);
            dataStyle.setBorderBottom(BorderStyle.THIN);
            dataStyle.setBorderLeft(BorderStyle.THIN);
            dataStyle.setBorderRight(BorderStyle.THIN);
            dataStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            dataStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            dataStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            dataStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            var headerRow = sheet.createRow(0);
            headerRow.setHeightInPoints(20);
            var ci = new AtomicInteger();
            doc.getHeaders().forEach(title -> {
                var cell = headerRow.createCell(ci.getAndIncrement(), CellType.STRING);
                cell.setCellValue(title);
                cell.setCellStyle(headerStyle);
            });

            var ri = new AtomicInteger(1);
            for (var rowValues : doc.getRows()) {
                var row = sheet.createRow(ri.getAndIncrement());
                var cj = new AtomicInteger();
                rowValues.forEach(raw -> {
                    var cell = row.createCell(cj.getAndIncrement());
                    var val = converter.convert(raw, raw != null ? raw.getClass() : String.class);
                    switch (val) {
                        case Number n -> cell.setCellValue(n.doubleValue());
                        case Boolean b -> cell.setCellValue(b);
                        case null -> cell.setCellValue("");
                        default -> cell.setCellValue(val.toString());
                    }
                    cell.setCellStyle(dataStyle);
                });
            }

            IntStream.range(0, doc.getHeaders().size()).forEach(c -> {
                if (doc.getColumnWidths().containsKey(c)) {
                    sheet.setColumnWidth(c, doc.getColumnWidths().get(c));
                } else {
                    sheet.autoSizeColumn(c);
                }
            });

            doc.getRowHeights().forEach((r, h) -> {
                var row = sheet.getRow(r);
                if (row != null) row.setHeight(h);
            });

            workbook.write(out);
        } catch (Exception e) {
            throw new RuntimeException("Failed to write Excel file", e);
        }
    }
}
