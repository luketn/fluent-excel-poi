package com.mycodefu.poi.examples.fluentpoi;

import com.github.javafaker.Faker;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Instant;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.Date;

import static java.time.temporal.ChronoUnit.DAYS;
import static org.junit.jupiter.api.Assertions.assertEquals;

class CellTest {
    @Test
    public void testOriginal() throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFDataFormat dataFormat = wb.createDataFormat();
            short df = dataFormat.getFormat("dd/mm/yyyy");
            XSSFCellStyle dateCellStyle = wb.createCellStyle();
            dateCellStyle.setDataFormat(df);

            XSSFCell cell = wb.createSheet("Test").createRow(0).createCell(0);
            cell.setCellStyle(dateCellStyle);
            Instant instant = Instant.from(ZonedDateTime.of(2020, 9, 27, 0, 0, 0, 0, ZoneId.systemDefault()));
            Date date = Date.from(instant);
            cell.setCellValue(date);

            try (FileOutputStream stream = new FileOutputStream(new File("output/rawpoi.xlsx"))) {
                wb.write(stream);
            }

            //check that the date value was written to the cell
            XSSFSheet worksheet = wb.getSheet("Test");
            testDateInCell(worksheet, 0, 0, instant, "27/09/2020");
        }
    }

    @Test
    public void testCellValue() {
        Instant date1 = Instant.from(ZonedDateTime.of(2020, 9, 27, 0, 0, 0, 0, ZoneId.systemDefault()));
        Instant date2 = date1.plus(1, DAYS);

        Book book = Book.create()
                .sheet("Explore")
                .value(0, 0, date1)
                .cell(0, 1).dateFormat("dd-mmm-yy").value(date2).end().end()
                .value(0, 2, "hi there")
                .done();

        book.write("output/fluentcell.xlsx");

        testDateInCell(book.sheet("Explore").worksheet, 0,0, date1, "27/09/2020");
        testDateInCell(book.sheet("Explore").worksheet, 0,1, date2, "28-Sep-20");
    }

    @Test
    public void testManyRows() {
        Sheet explore = Book.create().sheet("Explore");
        explore.value(0,0, "Name");
        explore.value(0,1, "Job");
        explore.value(0,2, "Hired");

        Instant startDate = Instant.from(ZonedDateTime.of(2020, 9, 27, 0, 0, 0, 0, ZoneId.systemDefault()));
        for(int i=1; i <= 100; i++) {
            Row row = explore.row(i);
            row.cell(0).value(Faker.instance().name().fullName());
            row.cell(1).value(Faker.instance().job().title());
            row.cell(2).value(startDate.plus(i * 7, DAYS));
        }

        explore.done().write("output/fluentmanyrows.xlsx");
    }



    private void testDateInCell(XSSFSheet worksheet, int row, int column, Instant instant, String expectedStringDate) {
        XSSFCell cellToTest = worksheet.getRow(row).getCell(column);
        assertEquals(instant, cellToTest.getDateCellValue().toInstant());

        //check that the string value of the cell is as expected - by converting the cell type to string
        DataFormatter formatter = new DataFormatter();
        String actualStringDate = formatter.formatCellValue(cellToTest);

        System.out.printf("Expected cell string value at %d, %d: %s\n", row, column, actualStringDate);
        assertEquals(expectedStringDate, actualStringDate);
    }
}