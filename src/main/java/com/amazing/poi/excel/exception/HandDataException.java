package com.amazing.poi.excel.exception;

/**
 * @author amzing.
 * @version 0.1
 * @time 2019/7/15 23:23
 * @since 0.1
 */
public class HandDataException extends RuntimeException {

    public HandDataException() {
        super();
    }

    public HandDataException(String message) {
        super(message);
    }

    public HandDataException(String message, Throwable cause) {
        super(message, cause);
    }

    public HandDataException(Throwable cause) {
        super(cause);
    }

    protected HandDataException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
