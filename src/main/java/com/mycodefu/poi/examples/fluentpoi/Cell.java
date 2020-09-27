package com.mycodefu.poi.examples.fluentpoi;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.time.Instant;
import java.util.Date;

public class Cell {
    public static final String DEFAULT_DATE_FORMAT = "dd/mm/yyyy";
    private final Book book;
    private final Sheet sheet;
    private final Row row;
    private final XSSFCell workcell;
    private String dateFormat;
    private boolean bold;
    private Object value;

    private Cell(Book book, Sheet sheet, Row row, XSSFCell workcell) {
        this.book = book;
        this.sheet = sheet;
        this.row = row;
        this.workcell = workcell;

        this.dateFormat = null;
        this.bold = false;
        this.value = null;
    }

    public static Cell create(Book book, Sheet sheet, Row row, int column) {
        XSSFCell workcell = row.workrow.getCell(column);
        if (workcell == null) {
            workcell = row.workrow.createCell(column);
        }
        return new Cell(book, sheet, row, workcell);
    }

    public Cell value(String value) {
        workcell.setCellValue(value);
        workcell.setCellType(CellType.STRING);
        this.value = value;
        setCellStyles();
        return this;
    }

    public Cell value(Instant value) {
        workcell.setCellValue(Date.from(value));
        this.value = value;
        setCellStyles();
        return this;
    }

    public Cell bold() {
        this.bold = true;
        return this;
    }

    public Cell dateFormat(String dateFormat) {
        this.dateFormat = dateFormat;
        return this;
    }

    public Row end() {
        return row;
    }

    public Book done() {
        return book;
    }

    private void setCellStyles() {
        if (this.value != null) {
            if (bold || dateFormat != null || this.value instanceof Instant) {
                XSSFCellStyle cellStyle = book.workbook.createCellStyle();

                if (value instanceof Instant) {
                    short df;
                    if (dateFormat != null) {
                        df = book.workbook.createDataFormat().getFormat(dateFormat);
                    } else {
                        df = book.workbook.createDataFormat().getFormat(DEFAULT_DATE_FORMAT);
                    }
                    cellStyle.setDataFormat(df);
                }

                if (bold) {
                    XSSFFont font = book.workbook.createFont();
                    font.setBold(true);
                    cellStyle.setFont(font);
                }

                this.workcell.setCellStyle(cellStyle);
            }
        }
    }
}
