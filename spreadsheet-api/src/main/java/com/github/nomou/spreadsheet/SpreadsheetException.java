package com.github.nomou.spreadsheet;

/**
 * Spreadsheet exception.
 *
 * @author vacoor
 * @since 1.0
 */
public class SpreadsheetException extends RuntimeException {
    public SpreadsheetException() {
    }

    public SpreadsheetException(final String message) {
        super(message);
    }

    public SpreadsheetException(final String message, final Throwable cause) {
        super(message, cause);
    }

    public SpreadsheetException(final Throwable cause) {
        super(cause);
    }

}
