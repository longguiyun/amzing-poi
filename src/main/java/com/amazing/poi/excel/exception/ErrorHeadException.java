package com.amazing.poi.excel.exception;

/**
 * <p>错误的表头</p>
 *
 * @author amzing.
 * @version 0.1
 * @time 2019/10/22 23:30
 * @since 0.1
 */
public class ErrorHeadException extends HandDataException {

    public ErrorHeadException() {
    }

    public ErrorHeadException(String message) {
        super(message);
    }

    public ErrorHeadException(String message, Throwable cause) {
        super(message, cause);
    }

    public ErrorHeadException(Throwable cause) {
        super(cause);
    }

    public ErrorHeadException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
