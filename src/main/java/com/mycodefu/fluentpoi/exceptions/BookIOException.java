package com.mycodefu.fluentpoi.exceptions;

import java.io.IOException;

public class BookIOException extends RuntimeException {
    public BookIOException(String message, IOException e) {
        super(message, e);
    }
}
