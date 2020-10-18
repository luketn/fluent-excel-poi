package com.mycodefu.poi.examples.fluentpoi;

import com.mycodefu.poi.examples.fluentpoi.exceptions.BookFileNotFoundException;
import com.mycodefu.poi.examples.fluentpoi.exceptions.BookIOException;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

/**
 * A fluent interface for writing a simple spreadsheet with basic styles.
 *
 * Note: Exceptions thrown by these classes extend RuntimeException, making their handling optional.
 * Specifically thrown exceptions will be documented in JavaDoc.
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

    /**
     * Open an Excel workbook and return a Fluent wrapper Book instance to interact with it.
     * @param filePath The file path to open.
     * @return An instance of Book.
     *
     * @throws BookFileNotFoundException when the file is not found.
     * @throws BookIOException when there is an I/O error reading the file.
     */
    public static Book open(String filePath) {
        File file = new File(filePath);
        if (!file.exists()) {
            throw new BookFileNotFoundException(String.format("File not found '%s'.", filePath));
        }
        try (FileInputStream fileInputStream = new FileInputStream(file)){
            return open(fileInputStream);
        } catch (IOException e) {
            throw new BookIOException("Failed to read the workbook.", e);
        }
    }

    /**
     * Open an Excel workbook from an InputStream.
     *
     * @param inputStream to read from.
     * @return An instance of Book.
     *
     * @throws BookIOException when there is an I/O error reading the stream.
     */
    public static Book open(InputStream inputStream) {
        try {
            return new Book(new XSSFWorkbook(inputStream));
        } catch (IOException e) {
            throw new BookIOException("Failed to read the workbook.", e);
        }
    }

    public Sheet sheet(String name) {
        return Sheet.create(this, name);
    }

    public void write(File file) {
        try (FileOutputStream fileOutputStream = new FileOutputStream(file)) {
            write(fileOutputStream);
        } catch (FileNotFoundException e) {
            throw new BookFileNotFoundException("Failed to write the workbook. File not found.", e);
        } catch (IOException e) {
            throw new BookIOException("Failed to write the workbook.", e);
        }
    }

    public void write(OutputStream stream) {
        try {
            this.workbook.write(stream);
        } catch (IOException e) {
            throw new BookIOException("Failed to write the workbook to the output stream.", e);
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
            throw new BookIOException("Failed to close the workbook.", e);
        }
    }
}
