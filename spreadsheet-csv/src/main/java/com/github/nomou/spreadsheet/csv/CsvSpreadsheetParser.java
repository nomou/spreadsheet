package com.github.nomou.spreadsheet.csv;

import com.github.nomou.spreadsheet.AbstractSpreadsheetParser;
import com.github.nomou.spreadsheet.SpreadsheetException;
import com.github.nomou.spreadsheet.SpreadsheetParser;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.util.NoSuchElementException;

/**
 * CSV spreadsheet parser.
 *
 * @author vacoor
 * @since 1.0
 */
class CsvSpreadsheetParser extends AbstractSpreadsheetParser {
    private static final int CSV_SHEETS = 1;
    private static final String CSV_SHEET_NAME = "sheet1";

    private final InputStream in;
    private CsvOptions options;
    private CsvParser parser;
    private String[] cells;
    private int row = -1;
    private int col = -1;
    private String value;

    public CsvSpreadsheetParser(final InputStream in) {
        this(in, CsvWriter.GB2312);
    }

    public CsvSpreadsheetParser(final InputStream in, final Charset encoding) {
        this.in = in;
        this.eventType = START_WORKBOOK;
        this.options = new CsvOptions(encoding, CsvLineParser.DEFAULT_SEPARATOR, CsvLineParser.DEFAULT_QUOTE_CHARACTER, CsvLineParser.DEFAULT_ESCAPE_CHARACTER);
    }

    @Override
    public SpreadsheetParser configure(final String option, final Object value) {
        options.set(option, value);
        return this;
    }

    @Override
    public int getNumberOfWorksheets() {
        return CSV_SHEETS;
    }

    @Override
    protected int doNext() throws SpreadsheetException {
        final CsvParser csvParser = getInternalParser();
        final int event = this.eventType;
        int newEvent = EOF;
        try {
            if (START_WORKBOOK == event) {
                this.worksheetIndex = 0;
                this.worksheetName = CSV_SHEET_NAME;
                newEvent = START_WORKSHEET;
            } else if (START_WORKSHEET == event || END_RECORD == event) {
                if (csvParser.hasNext()) {
                    this.cells = csvParser.next();
                    this.cells = null != this.cells ? this.cells : new String[0];
                    this.col = 0;
                    this.row = csvParser.getLineNumber();
                    newEvent = START_RECORD;
                }
                if (1 > this.cells.length) {
                    // close
                    this.close();
                    newEvent = END_WORKSHEET;
                }
            } else if (START_RECORD == event) {
                // this.eventType
                if (col >= cells.length) {
                    throw new NoSuchElementException();
                }
                value = this.cells[col];
                newEvent = START_CELL;
            } else if (START_CELL == event) {
                newEvent = END_CELL;
            } else if (END_CELL == event) {
                col++;
                if (col == cells.length) {
                    col = -1;
                    newEvent = END_RECORD;
                } else if (col < cells.length) {
                    value = cells[col];
                    newEvent = START_CELL;
                } else {
                    throw new NoSuchElementException();
                }
            } else if (END_WORKSHEET == event) {
                newEvent = END_WORKBOOK;
            }
        } catch (final IOException e) {
            throw new SpreadsheetException(e);
        }
        return newEvent;
    }

    private CsvParser getInternalParser() {
        if (null == parser) {
            parser = options.createParser(in);
        }
        return parser;
    }

    @Override
    public int getRow() {
        return row;
    }

    @Override
    public int getCol() {
        return col;
    }

    @Override
    public Object getValue() {
        final int st = this.eventType;
        if (START_CELL != st && END_CELL != st) {
            throw new IllegalStateException("getValue() called in illegal state");
        }
        return this.value;
    }

    @Override
    public void close() throws SpreadsheetException {
        try {
            if (null != this.parser) {
                this.parser.close();
            }
        } catch (final IOException e) {
            throw new SpreadsheetException(e);
        }
    }
}
