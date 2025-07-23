package com.excelninja.infrastructure.io;

import com.excelninja.domain.exception.DocumentConversionException;
import com.excelninja.domain.exception.HeaderMismatchException;
import com.excelninja.domain.exception.InvalidDocumentStructureException;
import com.excelninja.domain.model.ExcelSheet;
import com.excelninja.domain.model.ExcelWorkbook;
import com.excelninja.domain.model.Headers;
import com.excelninja.domain.port.WorkbookReader;
import com.excelninja.infrastructure.converter.DefaultConverter;
import com.excelninja.infrastructure.metadata.EntityMetadata;
import com.excelninja.infrastructure.metadata.FieldMapping;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.*;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.logging.Level;
import java.util.logging.Logger;

public class StreamingWorkbookReader implements WorkbookReader {
    private static final Logger logger = Logger.getLogger(StreamingWorkbookReader.class.getName());

    @Override
    public ExcelWorkbook read(File excelFile) throws IOException {
        try (FileInputStream fis = new FileInputStream(excelFile)) {
            return read(fis);
        }
    }

    @Override
    public ExcelWorkbook read(InputStream inputStream) throws IOException {
        try (OPCPackage opcPackage = OPCPackage.open(inputStream)) {
            return readFromOPCPackage(opcPackage);
        } catch (Exception e) {
            throw new DocumentConversionException("Failed to read Excel file with streaming reader", e);
        }
    }

    private ExcelWorkbook readFromOPCPackage(OPCPackage opcPackage) throws Exception {
        XSSFReader xssfReader = new XSSFReader(opcPackage);
        SharedStringsTable sharedStringsTable = (SharedStringsTable) xssfReader.getSharedStringsTable();
        StylesTable stylesTable = xssfReader.getStylesTable();
        Map<String, ExcelSheet> sheets = new LinkedHashMap<>();
        XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator) xssfReader.getSheetsData();

        while (sheetIterator.hasNext()) {
            try (InputStream sheetStream = sheetIterator.next()) {
                String sheetName = sheetIterator.getSheetName();
                ExcelSheet sheet = readSheetWithStreaming(sheetStream, sheetName, sharedStringsTable, stylesTable);
                sheets.put(sheetName, sheet);
            }
        }
        if (sheets.isEmpty()) throw new InvalidDocumentStructureException("No sheets found in workbook");
        ExcelWorkbook.WorkbookBuilder builder = ExcelWorkbook.builder();
        sheets.forEach(builder::sheet);
        return builder.build();
    }

    private ExcelSheet readSheetWithStreaming(
            InputStream sheetStream,
            String sheetName,
            SharedStringsTable sst,
            StylesTable styles
    ) throws Exception {
        SheetAndHeaderHandler handler = new SheetAndHeaderHandler(sst, styles);
        XMLReader xmlReader = XMLHelper.newXMLReader();
        xmlReader.setContentHandler(handler);
        xmlReader.parse(new InputSource(sheetStream));
        return handler.buildExcelSheet(sheetName);
    }

    public <T> Iterator<List<T>> readInChunks(
            File file,
            Class<T> entityType,
            int chunkSize
    ) throws IOException {
        return new ChunkIterator<>(Files.newInputStream(file.toPath()), entityType, chunkSize, true);
    }

    public <T> Iterator<List<T>> readInChunks(
            InputStream inputStream,
            Class<T> entityType,
            int chunkSize
    ) {
        return new ChunkIterator<>(inputStream, entityType, chunkSize, false);
    }

    private static class BaseSheetHandler extends DefaultHandler {
        protected final SharedStringsTable sst;
        protected final StylesTable stylesTable;
        private String currentCellRef;
        private String currentCellType;
        private int currentCellStyleIndex;
        private final StringBuilder currentCellValue = new StringBuilder();

        protected Map<Integer, Object> currentRowData;

        public BaseSheetHandler(
                SharedStringsTable sst,
                StylesTable styles
        ) {
            this.sst = sst;
            this.stylesTable = styles;
        }

        @Override
        public void startElement(
                String uri,
                String localName,
                String qName,
                Attributes attributes
        ) {
            if ("row".equals(qName)) {
                currentRowData = new HashMap<>();
            } else if ("c".equals(qName)) {
                currentCellRef = attributes.getValue("r");
                currentCellType = attributes.getValue("t");
                String styleStr = attributes.getValue("s");
                currentCellStyleIndex = styleStr != null ? Integer.parseInt(styleStr) : -1;
                currentCellValue.setLength(0);
            }
        }

        @Override
        public void characters(
                char[] ch,
                int start,
                int length
        ) {
            currentCellValue.append(ch, start, length);
        }

        @Override
        public void endElement(
                String uri,
                String localName,
                String qName
        ) {
            if ("c".equals(qName)) {
                int colIdx = CellReference.convertColStringToIndex(currentCellRef.replaceAll("\\d", ""));
                Object value = parseValue(currentCellValue.toString(), currentCellType, currentCellStyleIndex, sst, stylesTable);
                currentRowData.put(colIdx, value);
            } else if ("row".equals(qName)) {
                processRow();
            }
        }

        protected void processRow() { /* To be implemented by subclasses */ }
    }

    private static class SheetAndHeaderHandler extends BaseSheetHandler {
        private final List<String> headers = new ArrayList<>();
        private final List<List<Object>> allRows = new ArrayList<>();
        private boolean isHeaderProcessed = false;
        private int maxColCount = 0;

        public SheetAndHeaderHandler(
                SharedStringsTable sst,
                StylesTable styles
        ) {
            super(sst, styles);
        }

        @Override
        protected void processRow() {
            if (currentRowData.isEmpty()) return;
            maxColCount = Math.max(maxColCount, currentRowData.keySet().stream().max(Integer::compareTo).orElse(-1) + 1);
            List<Object> rowValues = new ArrayList<>(Collections.nCopies(maxColCount, null));
            currentRowData.forEach(rowValues::set);

            if (!isHeaderProcessed) {
                rowValues.forEach(val -> headers.add(val != null ? val.toString().trim() : ""));
                isHeaderProcessed = true;
            } else {
                allRows.add(rowValues);
            }
        }

        public ExcelSheet buildExcelSheet(String sheetName) {
            if (headers.isEmpty())
                throw new InvalidDocumentStructureException("No headers found in sheet: " + sheetName);
            return ExcelSheet.builder().name(sheetName).headers(headers).rows(allRows).build();
        }
    }

    private static class ChunkIterator<T> implements Iterator<List<T>>, AutoCloseable {
        private final int chunkSize;
        private final BlockingQueue<Object> queue = new LinkedBlockingQueue<>(20000); // Buffer size
        private final InputStream managedInputStream;
        private final boolean closeOnFinish;
        private final Thread producerThread;

        private List<T> nextChunk;
        private volatile boolean isProducerFinished = false;
        private volatile Exception producerException = null;

        private static final Object END_OF_QUEUE = new Object(); // Poison Pill

        public ChunkIterator(
                InputStream inputStream,
                Class<T> entityType,
                int chunkSize,
                boolean closeOnFinish
        ) {
            this.chunkSize = chunkSize;
            this.managedInputStream = inputStream;
            this.closeOnFinish = closeOnFinish;

            // 생산자 스레드 시작
            this.producerThread = new Thread(() -> {
                try (OPCPackage opcPackage = OPCPackage.open(managedInputStream)) {
                    XSSFReader xssfReader = new XSSFReader(opcPackage);
                    SharedStringsTable sst = (SharedStringsTable) xssfReader.getSharedStringsTable();
                    StylesTable styles = xssfReader.getStylesTable();

                    XMLReader xmlReader = XMLHelper.newXMLReader();
                    xmlReader.setContentHandler(new ChunkingHandler(entityType, sst, styles));

                    try (InputStream sheetStream = xssfReader.getSheetsData().next()) {
                        xmlReader.parse(new InputSource(sheetStream));
                    }
                } catch (Exception e) {
                    producerException = e;
                } finally {
                    try {
                        queue.put(END_OF_QUEUE);
                    } catch (InterruptedException e) {
                        Thread.currentThread().interrupt();
                    }
                    isProducerFinished = true;
                }
            });
            this.producerThread.start();
        }

        @Override
        public boolean hasNext() {
            if (producerException != null) {
                throw new DocumentConversionException("Error in background producer thread", producerException);
            }
            if (nextChunk != null && !nextChunk.isEmpty()) {
                return true;
            }
            fillNextChunk();
            return nextChunk != null && !nextChunk.isEmpty();
        }

        @Override
        public List<T> next() {
            if (!hasNext()) throw new NoSuchElementException("No more chunks available.");
            List<T> chunkToReturn = nextChunk;
            nextChunk = null;
            return chunkToReturn;
        }

        private void fillNextChunk() {
            if (isProducerFinished && queue.isEmpty()) {
                nextChunk = Collections.emptyList();
                return;
            }

            nextChunk = new ArrayList<>(chunkSize);
            while (nextChunk.size() < chunkSize) {
                try {
                    Object item = queue.poll(); // Non-blocking poll
                    if (item == null) { // Queue is empty, but producer may not be finished
                        if (isProducerFinished && queue.isEmpty()) break;
                        item = queue.take(); // Block until an item is available
                    }

                    if (item == END_OF_QUEUE) {
                        isProducerFinished = true;
                        queue.offer(END_OF_QUEUE); // Put it back for other consumers if any
                        break;
                    }
                    nextChunk.add((T) item);
                } catch (InterruptedException e) {
                    Thread.currentThread().interrupt();
                    throw new DocumentConversionException("Iterator thread interrupted", e);
                }
            }
        }

        @Override
        public void close() {
            producerThread.interrupt();
            if (closeOnFinish && managedInputStream != null) {
                try {
                    managedInputStream.close();
                } catch (IOException e) {
                    logger.log(Level.WARNING, "Error closing input stream", e);
                }
            }
        }

        private class ChunkingHandler extends BaseSheetHandler {
            private final EntityMetadata<T> entityMetadata;
            private final DefaultConverter converter = new DefaultConverter();
            private boolean isHeaderProcessed = false;
            private List<Integer> fieldToColumnMapping;
            private int maxColCount = 0;

            public ChunkingHandler(
                    Class<T> entityType,
                    SharedStringsTable sst,
                    StylesTable styles
            ) {
                super(sst, styles);
                this.entityMetadata = EntityMetadata.of(entityType);
            }

            @Override
            protected void processRow() {
                if (currentRowData.isEmpty()) return;
                maxColCount = Math.max(maxColCount, currentRowData.keySet().stream().max(Integer::compareTo).orElse(-1) + 1);
                List<Object> rowValues = new ArrayList<>(Collections.nCopies(maxColCount, null));
                currentRowData.forEach(rowValues::set);

                if (!isHeaderProcessed) {
                    List<String> headers = new ArrayList<>();
                    rowValues.forEach(val -> headers.add(val != null ? val.toString().trim() : ""));
                    prepareFieldMapping(headers);
                    isHeaderProcessed = true;
                } else {
                    try {
                        queue.put(convertRowToEntity(rowValues));
                    } catch (Exception e) {
                        logger.warning("Failed to process or queue row: " + e.getMessage());
                    }
                }
            }

            private void prepareFieldMapping(List<String> headers) {
                Headers sheetHeaders = Headers.of(headers);
                fieldToColumnMapping = new ArrayList<>();
                for (FieldMapping fieldMapping : entityMetadata.getReadFieldMappings()) {
                    String headerName = fieldMapping.getHeaderName();
                    if (!sheetHeaders.containsHeader(headerName)) {
                        throw new HeaderMismatchException("Header not found: " + headerName, "missing");
                    }
                    fieldToColumnMapping.add(sheetHeaders.getPositionOf(headerName));
                }
            }

            private T convertRowToEntity(List<Object> rowValues) throws Exception {
                T entity = entityMetadata.createInstance();
                List<FieldMapping> fieldMappings = entityMetadata.getReadFieldMappings();
                for (int i = 0; i < fieldMappings.size(); i++) {
                    FieldMapping fieldMapping = fieldMappings.get(i);
                    int colIdx = fieldToColumnMapping.get(i);
                    Object cellValue = (colIdx < rowValues.size()) ? rowValues.get(colIdx) : null;
                    fieldMapping.setValue(entity, cellValue, converter);
                }
                return entity;
            }
        }
    }

    private static Object parseValue(
            String value,
            String type,
            int styleIndex,
            SharedStringsTable sst,
            StylesTable styles
    ) {
        if (value == null || value.isEmpty()) return null;
        String cellType = (type != null) ? type : "";
        try {
            switch (cellType) {
                case "s":
                    return sst.getItemAt(Integer.parseInt(value)).getString();
                case "str":
                    return value;
                case "b":
                    return "1".equals(value);
                case "e":
                    return "ERROR: " + value;
                default:
                    double d = Double.parseDouble(value);
                    if (styleIndex != -1) {
                        XSSFCellStyle style = styles.getStyleAt(styleIndex);
                        if (DateUtil.isADateFormat(style.getDataFormat(), style.getDataFormatString()) && DateUtil.isValidExcelDate(d)) {
                            return DateUtil.getJavaDate(d);
                        }
                    }
                    if (d == Math.floor(d) && !Double.isInfinite(d) && d <= Long.MAX_VALUE) {
                        return (long) d;
                    }
                    return d;
            }
        } catch (NumberFormatException e) {
            return value;
        }
    }
}