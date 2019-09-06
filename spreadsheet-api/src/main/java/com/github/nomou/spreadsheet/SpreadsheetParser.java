package com.github.nomou.spreadsheet;

import java.util.NoSuchElementException;

/**
 * Spreadsheet parser definition.
 *
 * @author vacoor
 * @since 1.0
 */
public interface SpreadsheetParser extends SpreadsheetConfigurable<SpreadsheetParser> {
    /**
     * Workbook start event.
     */
    int START_WORKBOOK = 11;

    /**
     * Worksheet start event.
     */
    int START_WORKSHEET = 21;

    /**
     * Record(row) start event.
     */
    int START_RECORD = 31;

    /**
     * Cell start event.
     */
    int START_CELL = 41;

    /**
     * Cell end event.
     */
    int END_CELL = 42;

    /**
     * Record(row) end event.
     */
    int END_RECORD = 32;

    /**
     * Worksheet end event.
     */
    int END_WORKSHEET = 22;

    /**
     * Workbook end event.
     */
    int END_WORKBOOK = 12;

    /**
     * Stream end event.
     */
    int EOF = -1;


    /**
     * Returns an integer code that indicates the type of the event the cursor is pointing to.
     *
     * @return the parsing event type
     */
    int getEventType();


    /* ********************************** *
     *             Worksheets
     * ********************************** */

    /**
     * Returns the number of worksheets in the current workbook.
     *
     * @return number of worksheets.
     */
    int getNumberOfWorksheets();

    /**
     * Returns the current worksheet index in workbook of the parse event.
     *
     * @return the current worksheet index
     * @throws java.lang.IllegalStateException if this state is not a valid state.
     */
    int getWorksheetIndex();

    /**
     * Returns the current worksheet name of the parse event.
     *
     * @return the current worksheet name
     * @throws java.lang.IllegalStateException if this state is not a valid state.
     */
    String getWorksheetName();

    /**
     * Returns the current row index of the parse event.
     *
     * @return the current row index
     * @throws java.lang.IllegalStateException if this state is not a valid row state.
     */
    int getRow();

    /**
     * Returns the current cell column index of the parse event.
     *
     * @return the current cell column index
     * @throws java.lang.IllegalStateException if this state is not a valid cell state.
     */
    int getCol();

    /**
     * Returns the current cell value of the parse event.
     *
     * @return the current cell or null
     * @throws java.lang.IllegalStateException if this state is not a valid cell state.
     */
    Object getValue();

    /**
     * Returns true if there are more parsing events and false if there are no more events.
     * This method will return false if the current state of the SpreadsheetParser is END_WORKBOOK.
     *
     * @return true if there are more events, false otherwise
     * @throws SpreadsheetException if there is a fatal error detecting the next state
     */
    boolean hasNext() throws SpreadsheetException;

    /**
     * Get next parsing event.
     * <p>
     * <p>Given the following workbook:<br>
     * sheet1: <br>
     * |----|----| <br>
     * | A1 | B1 | <br>
     * |----|----| <br>
     * The behavior of calling next() when being on foo will be:<br>
     * 1- the start worksheet(START_WORKSHEET)<br>
     * 3- then the start first row(START_RECORD)<br>
     * 4- then the start A1 cell(START_CELL)<br>
     * 5- then the end A1 cell(END_CELL)<br>
     * 6- then the start B1 cell(START_CELL)<br>
     * 7- then the end B1 cell(END_CELL)<br>
     * 8- then the end first row(END_RECORD)<br>
     * 9- then the end worksheet(END_WORKSHEET)<br>
     * 10- then the end workbook(END_WORKBOOK)<br>
     * <p>
     * This method will throw an IllegalStateException if it is called after hasNext() returns false.
     *
     * @return the integer code corresponding to the current parse event
     * @throws NoSuchElementException if this is called when hasNext() returns false
     * @throws SpreadsheetException   if there is an error processing the underlying input source
     */
    int next() throws SpreadsheetException;

    /**
     * Get next record(row).
     *
     * @param ignoreEmptyRecord skip empty record
     * @return null if there not more record, cell array otherwise
     * @throws SpreadsheetException if there is an error processing the underlying input source
     */
    Object[] nextRecord(final boolean ignoreEmptyRecord) throws SpreadsheetException;

    /**
     * Frees any resources associated with this Reader.  This method does not close the underlying input source.
     *
     * @throws SpreadsheetException if there are errors freeing associated resources
     */
    void close() throws SpreadsheetException;


}
