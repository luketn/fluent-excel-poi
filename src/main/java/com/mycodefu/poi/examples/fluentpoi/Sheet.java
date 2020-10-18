package com.mycodefu.poi.examples.fluentpoi;

import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.time.Instant;

public class Sheet {
    protected final Book book;
    protected final XSSFSheet worksheet;

    private Sheet(Book book, XSSFSheet worksheet) {
        this.book = book;
        this.worksheet = worksheet;
    }

    public static Sheet create(Book book, String name) {
        XSSFSheet worksheet = book.workbook.getSheet(name);
        if (worksheet==null){
            worksheet = book.workbook.createSheet(name);
        }
        return new Sheet(book, worksheet);
    }

    public Sheet erase() {
        while (worksheet.getPhysicalNumberOfRows() > 0) {
            worksheet.removeRow(worksheet.getRow(worksheet.getLastRowNum()));
        }
        return this;
    }

    public int rowCount() {
        return worksheet.getPhysicalNumberOfRows();
    }

    public Sheet autosizeColumn(int column) {
        worksheet.autoSizeColumn(column);
        return this;
    }

    public Row row(int row) {
        return Row.create(book, this, row);
    }

    public Cell cell(int row, int column) {
        return row(row).cell(column);
    }

    public Sheet setValue(int row, int column, String value) {
        return row(row).cell(column).setValue(value).end().end();
    }

    public Sheet setValue(int row, int column, Instant value) {
        return row(row).cell(column).setValue(value).end().end();
    }

    public Book done() {
        return book;
    }

}
