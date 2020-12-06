package com.mycodefu.fluentpoi;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
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
    private boolean currency;

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
        setCellStyles();
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
        return this;
    }

    public Cell currency() {
        this.currency = true;
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
            if (bold || currency || dateFormat != null || this.value instanceof Instant) {
                String styleKey = String.format("%b-%s-%b", bold, dateFormat, currency);
                if (!book.styles.containsKey(styleKey)) {
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

                    if (currency && value instanceof Double) {
                        XSSFDataFormat df = book.workbook.createDataFormat();
                        cellStyle.setDataFormat(df.getFormat("$#,##0.00"));
                    }

                    book.styles.put(styleKey, cellStyle);
                }

                this.workcell.setCellStyle(book.styles.get(styleKey));
            }
        }
    }

}
