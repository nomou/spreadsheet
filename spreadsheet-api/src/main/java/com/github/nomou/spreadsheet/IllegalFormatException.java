package com.github.nomou.spreadsheet;

/**
 * Illegal spreadsheet format exception.
 *
 * @author vacoor
 * @since 1.0
 */
public class IllegalFormatException extends SpreadsheetException {
    public IllegalFormatException() {
        super();
    }

    public IllegalFormatException(final String message) {
        super(message);
    }

    public IllegalFormatException(final String message, final Throwable cause) {
        super(message, cause);
    }

    public IllegalFormatException(final Throwable cause) {
        super(cause);
    }
}
