package com.mycodefu.poi.examples.fluentpoi;

import org.apache.poi.xssf.usermodel.XSSFRow;

public class Row {
    protected final Book book;
    protected final Sheet sheet;
    protected final int rowNum;
    protected final XSSFRow workrow;

    public Row(Book book, Sheet sheet, int rowNum, XSSFRow row) {
        this.book = book;
        this.sheet = sheet;
        this.rowNum = rowNum;
        this.workrow = row;
    }

    public static Row create(Book book, Sheet sheet, int row){
        XSSFRow worksheetRow = sheet.worksheet.getRow(row);
        if (worksheetRow==null){
            worksheetRow = sheet.worksheet.createRow(row);
        }
        return new Row(book, sheet, row, worksheetRow);
    }

    public Cell cell(int column){ return Cell.create(book, sheet, this, column);}

    public Sheet end(){return sheet;}

    public Book done(){return book;}
}
