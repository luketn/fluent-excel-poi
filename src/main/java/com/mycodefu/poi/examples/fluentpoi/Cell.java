package com.mycodefu.poi.examples.fluentpoi;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.time.Instant;
import java.util.Date;

public class Cell {
    public static final String DEFAULT_DATE_FORMAT = "dd/mm/yyyy";
    private final Book book;
    private final Sheet sheet;
    private final Row row;
    private final XSSFCell workcell;

    private Cell(Book book, Sheet sheet, Row row, XSSFCell workcell) {
        this.book = book;
        this.sheet = sheet;
        this.row = row;
        this.workcell = workcell;
    }

    public static Cell create(Book book, Sheet sheet, Row row, int column) {
        String key = String.format("%s-%s", row.rowNum, column);
        if (!book.workcells.containsKey(key)) {
            book.workcells.put(key, row.workrow.createCell(column));
        }
        return new Cell(book, sheet, row, book.workcells.get(key));
    }

    public Cell value(String value) {
        workcell.setCellValue(value);
        return this;
    }

    public Cell value(Instant value) {
        workcell.setCellValue(Date.from(value));
        workcell.setCellStyle(dateStyle(DEFAULT_DATE_FORMAT));
        return this;
    }

    public Cell dateFormat(String dateFormat) {
        putDateFormat(dateFormat);
        return this;
    }

    private CellStyle dateStyle(String dateFormat) {
        if (!book.workstyles.containsKey("date")) {
            putDateFormat(dateFormat);
        }
        return book.workstyles.get("date");
    }

    private void putDateFormat(String dateFormat) {
        short df = book.workbook.createDataFormat().getFormat(dateFormat);
        XSSFCellStyle dateCellStyle = book.workbook.createCellStyle();
        dateCellStyle.setDataFormat(df);
        book.workstyles.put("date", dateCellStyle);
    }

    public Row end() {
        return row;
    }

    public Book done() {
        return book;
    }
}
