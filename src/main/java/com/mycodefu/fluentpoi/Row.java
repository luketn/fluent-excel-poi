package com.mycodefu.fluentpoi;

import org.apache.poi.xssf.usermodel.XSSFRow;

public class Row {
    protected final Book book;
    protected final Sheet sheet;
    protected final int row;
    protected final XSSFRow workrow;

    public Row(Book book, Sheet sheet, int row, XSSFRow workrow) {
        this.book = book;
        this.sheet = sheet;
        this.row = row;
        this.workrow = workrow;
    }

    public static Row create(Book book, Sheet sheet, int row){
        XSSFRow workRow = sheet.worksheet.getRow(row);
        if (workRow==null){
            workRow = sheet.worksheet.createRow(row);
        }
        return new Row(book, sheet, row, workRow);
    }

    public Cell cell(int column){ return Cell.create(book, sheet, this, column);}

    //TODO: Add overload to get column by letter

    public Sheet end(){return sheet;}

    public Book done(){return book;}
}
