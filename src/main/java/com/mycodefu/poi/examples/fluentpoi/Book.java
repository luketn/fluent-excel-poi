package com.mycodefu.poi.examples.fluentpoi;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

/**
 * A fluent interface for writing a simple spreadsheet with basic styles.
 */
public class Book implements AutoCloseable {
    protected final XSSFWorkbook workbook;
    protected final Map<String, XSSFCellStyle> styles = new HashMap<>();

    private Book(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public static Book create() {
        return new Book(new XSSFWorkbook());
    }

    public static Book open(String filePath) {
        try {
            return new Book(new XSSFWorkbook(filePath));
        } catch (IOException e) {
            throw new RuntimeException("Failed to read the workbook.", e);
        }
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

    @Override
    public void close() {
        try {
            this.workbook.close();
        } catch (IOException e) {
            throw new RuntimeException("Failed to close the workbook.", e);
        }
    }
}
