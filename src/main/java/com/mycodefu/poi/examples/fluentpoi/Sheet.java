package com.mycodefu.poi.examples.fluentpoi;

import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.time.Instant;
import java.util.Date;

/**
 * A fluent interface for writing a simple spreadsheet with basic styles.
 */
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

    public Row row(int row) {
        return Row.create(book, this, row);
    }

    public Cell cell(int row, int column) {
        return row(row).cell(column);
    }

    public Sheet value(int row, int column, String value) {
        return row(row).cell(column).value(value).end().end();
    }

    public Sheet value(int row, int column, Instant value) {
        return row(row).cell(column).value(value).end().end();
    }

    public Book done() {
        return book;
    }
}
