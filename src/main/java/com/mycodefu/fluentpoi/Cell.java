package com.mycodefu.fluentpoi;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.time.Instant;
import java.util.Date;

public class Cell {
    public static final String DEFAULT_DATE_FORMAT = "dd/mm/yyyy";
    public static final String DEFAULT_CURRENCY_FORMAT = "\"$\"#,##0.00";

    private final Book book;
    private final Sheet sheet;
    private final Row row;
    private final XSSFCell workcell;
    private String dataFormat;
    private Object value;
    private boolean bold;

    private Cell(Book book, Sheet sheet, Row row, XSSFCell workcell) {
        this.book = book;
        this.sheet = sheet;
        this.row = row;
        this.workcell = workcell;

        this.bold = false;
        this.dataFormat = null;
        this.value = null;
    }

    public static Cell create(Book book, Sheet sheet, Row row, int column) {
        XSSFCell workcell = row.workrow.getCell(column);
        if (workcell == null) {
            workcell = row.workrow.createCell(column);
        }
        return new Cell(book, sheet, row, workcell);
    }

    public Cell setValue(String value) {
        workcell.setCellValue(value);
        workcell.setCellType(CellType.STRING);
        this.value = value;
        setCellStyles();
        return this;
    }

    public Cell setValue(Instant value) {
        workcell.setCellValue(Date.from(value));
        this.value = value;
        if (this.dataFormat == null) {
            format(DEFAULT_DATE_FORMAT);
        } else {
            setCellStyles();
        }
        return this;
    }

    public Cell setValue(double value) {
        workcell.setCellValue(value);
        this.value = value;
        setCellStyles();
        return this;
    }

    public double getValueAsDouble() {
        return workcell.getNumericCellValue();
    }

    public String getValueAsString() {
        return workcell.getStringCellValue();
    }

    public Instant getValueAsInstant() {
        return workcell.getDateCellValue().toInstant();
    }

    public Cell bold() {
        this.bold = true;
        if (value != null) {
            setCellStyles();
        }
        return this;
    }

    public Cell currency() {
        return format(DEFAULT_CURRENCY_FORMAT);
    }
    public Cell format(String format) {
        this.dataFormat = format;
        if (value != null) {
            setCellStyles();
        }
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
            if (bold || dataFormat != null) {
                this.workcell.setCellStyle(book.cellStyles.getCellStyle(dataFormat, bold));
            }
        }
    }

}
