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
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static java.time.temporal.ChronoUnit.DAYS;
import static org.junit.jupiter.api.Assertions.*;

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

        testDateInCell(book.sheet("Explore").worksheet, 0, 0, date1, "27/09/2020");
        testDateInCell(book.sheet("Explore").worksheet, 0, 1, date2, "28-Sep-20");
    }

    @Test
    public void testCreateJobsSheet() {
        Sheet jobs = Book.create().sheet("Jobs");

        jobs.cell(0, 0).bold().value("Name");
        jobs.cell(0, 1).bold().value("Job");
        jobs.cell(0, 2).bold().value("Hired");

        Instant startDate = Instant.from(ZonedDateTime.of(2020, 9, 27, 0, 0, 0, 0, ZoneId.systemDefault()));
        List<Job> jobList = new ArrayList<>();
        for (int i = 1; i <= 100; i++) {
            Job job = new Job(
                    Faker.instance().name().fullName(),
                    Faker.instance().job().title(),
                    startDate.plus(i * 7, DAYS)
            );

            Row row = jobs.row(i);
            row.cell(0).value(job.name);
            row.cell(1).value(job.job);
            row.cell(2).value(job.hiredDate);

            jobList.add(job);
        }

        Book book = jobs.done();
        book.write("output/fluentmanyrows.xlsx");

        XSSFSheet workbookSheet = book.workbook.getSheet("Jobs");

        XSSFCell testNameHeaderCell = workbookSheet.getRow(0).getCell(0);
        assertEquals("Name", testNameHeaderCell.getStringCellValue());
        assertTrue(testNameHeaderCell.getCellStyle().getFont().getBold());

        XSSFCell testJobHeaderCell = workbookSheet.getRow(0).getCell(1);
        assertEquals("Job", testJobHeaderCell.getStringCellValue());
        assertTrue(testJobHeaderCell.getCellStyle().getFont().getBold());

        XSSFCell testHiredHeaderCell = workbookSheet.getRow(0).getCell(2);
        assertEquals("Hired", testHiredHeaderCell.getStringCellValue());
        assertTrue(testHiredHeaderCell.getCellStyle().getFont().getBold());

        for (int i = 0; i < jobList.size(); i++) {
            Job job = jobList.get(i);
            XSSFRow row = workbookSheet.getRow(i + 1);

            XSSFCell testNameCell = row.getCell(0);
            assertEquals(job.name, testNameCell.getStringCellValue());
            assertFalse(testNameCell.getCellStyle().getFont().getBold());

            XSSFCell testJobCell = row.getCell(1);
            assertEquals(job.job, testJobCell.getStringCellValue());
            assertFalse(testJobCell.getCellStyle().getFont().getBold());

            XSSFCell testHiredDateCell = row.getCell(2);
            DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("dd/MM/yyyy").withZone(ZoneId.systemDefault());
            String expectedDateString = dateTimeFormatter.format(job.hiredDate);
            testDateInCell(job.hiredDate, expectedDateString, testHiredDateCell);
            assertFalse(testHiredDateCell.getCellStyle().getFont().getBold());
        }

    }


    private void testDateInCell(XSSFSheet worksheet, int row, int column, Instant instant, String expectedStringDate) {
        XSSFCell cellToTest = worksheet.getRow(row).getCell(column);
        String actualDateString = testDateInCell(instant, expectedStringDate, cellToTest);
        System.out.printf("Expected cell string value at %d, %d: %s\n", row, column, actualDateString);
    }

    private String testDateInCell(Instant instant, String expectedStringDate, XSSFCell cellToTest) {
        assertEquals(instant, cellToTest.getDateCellValue().toInstant());

        //check that the string value of the cell is as expected - by converting the cell type to string
        DataFormatter formatter = new DataFormatter();
        String actualStringDate = formatter.formatCellValue(cellToTest);
        assertEquals(expectedStringDate, actualStringDate);
        return actualStringDate;
    }

    record Job(String name, String job, Instant hiredDate) { }
}