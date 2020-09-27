package com.mycodefu.poi.examples.fluentpoi;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

public class Book {
    protected final Map<String, CellStyle> workstyles = new HashMap<>();

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
