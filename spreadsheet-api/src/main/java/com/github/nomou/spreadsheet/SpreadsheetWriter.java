package com.github.nomou.spreadsheet;

import java.util.Date;

/**
 * Spreadsheet writer definition.
 *
 * @author vacoor
 * @since 1.0
 */
public interface SpreadsheetWriter extends SpreadsheetConfigurable<SpreadsheetWriter> {

    /**
     * Create a worksheet and start writing to it.
     *
     * @return the current spreadsheet writer
     * @throws SpreadsheetException if a write error occurs
     */
    SpreadsheetWriter start() throws SpreadsheetException;

    /**
     * Create a worksheet with the given name and start writing to it.
     *
     * @return the current spreadsheet writer
     * @throws SpreadsheetException if a write error occurs
     */
    SpreadsheetWriter start(final String worksheetName) throws SpreadsheetException;

    /**
     * Write the given Boolean value.
     *
     * @return the current spreadsheet writer
     * @throws SpreadsheetException if a write error occurs
     */
    SpreadsheetWriter write(final Boolean bool) throws SpreadsheetException;

    /**
     * Write the given number value.
     *
     * @param number the value
     * @return the current spreadsheet writer
     * @throws SpreadsheetException if a write error occurs
     */
    SpreadsheetWriter write(final Number number) throws SpreadsheetException;

    /**
     * Write the given date value using the default pattern.
     *
     * @param date the value
     * @return the current spreadsheet writer
     * @throws SpreadsheetException if a write error occurs
     */
    SpreadsheetWriter write(final Date date) throws SpreadsheetException;

    /**
     * Write the date value using the given pattern.
     *
     * @param date the value
     * @return the current spreadsheet writer
     * @throws SpreadsheetException if a write error occurs
     */
    SpreadsheetWriter write(final Date date, final String pattern) throws SpreadsheetException;

    /**
     * Write the text.
     *
     * @param text the text
     * @return the current spreadsheet writer
     * @throws SpreadsheetException if a write error occurs
     */
    SpreadsheetWriter write(final String text) throws SpreadsheetException;

    /**
     * Write the object.
     *
     * <p>
     * if 'obj' is a 'Boolean/Number/Date/Calendar/String' value or an array of them,
     * the corresponding 'write' method will be called for writing.</p>
     *
     * @param obj the object
     * @return the current spreadsheet writer
     * @throws SpreadsheetException if a write error occurs
     */
    SpreadsheetWriter write(final Object obj) throws SpreadsheetException;

    /**
     * Writes the given values.
     *
     * @param cells the values
     * @param <E>   the type of values
     * @return the current spreadsheet writer
     * @throws SpreadsheetException if a write error occurs
     */
    @SuppressWarnings("unchecked")
    <E> SpreadsheetWriter write(E... cells) throws SpreadsheetException;

    /**
     * End the writing of the current record(row) and start the next record.
     *
     * @return the current spreadsheet writer
     * @throws SpreadsheetException if a write error occurs
     */
    SpreadsheetWriter next() throws SpreadsheetException;

    /**
     * End the writing of the current spreadsheet and close it.
     *
     * @return the current spreadsheet writer
     * @throws SpreadsheetException if a write error occurs
     */
    void close() throws SpreadsheetException;

}
