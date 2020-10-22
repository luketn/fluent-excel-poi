package com.mycodefu.fluentpoi.exceptions;

import java.io.FileNotFoundException;

public class BookFileNotFoundException extends RuntimeException {
    public BookFileNotFoundException(String message) {
        super(message);
    }

    public BookFileNotFoundException(String message, FileNotFoundException e) {
        super(message, e);
    }
}
