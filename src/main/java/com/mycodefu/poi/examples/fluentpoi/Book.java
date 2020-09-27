package com.mycodefu.poi.examples.fluentpoi;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/**
 * A fluent interface for writing a simple spreadsheet with basic styles.
 */
public class Book {
    protected final XSSFWorkbook workbook;

    private Book() {
        this.workbook = new XSSFWorkbook();
    }

    public static Book create() {
        return new Book();
    }

    public Sheet sheet(String name) {
        return Sheet.create(this, name);
    }

    public void write(File file) {
        try (FileOutputStream fileOutputStream = new FileOutputStream(file)) {
            write(fileOutputStream);
        } catch (FileNotFoundException e) {
            throw new RuntimeException("Failed to write the workbook. File not found.", e);
        } catch (IOException e) {
            throw new RuntimeException("Failed to write the workbook.", e);
        }
    }

    public void write(OutputStream stream) {
        try {
            this.workbook.write(stream);
        } catch (IOException e) {
            throw new RuntimeException("Failed to write the workbook to the output stream.", e);
        }
    }

    public void write(String filePath) {
        write(new File(filePath));
    }
}
