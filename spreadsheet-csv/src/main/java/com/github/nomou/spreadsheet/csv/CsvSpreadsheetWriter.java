package com.github.nomou.spreadsheet.csv;

import com.github.nomou.spreadsheet.AbstractSpreadsheetWriter;
import com.github.nomou.spreadsheet.SpreadsheetException;
import com.github.nomou.spreadsheet.SpreadsheetWriter;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.text.FieldPosition;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

/**
 * CSV spreadsheet writer.
 *
 * @author vacoor
 * @since 1.0
 */
class CsvSpreadsheetWriter extends AbstractSpreadsheetWriter {
    private final OutputStream out;
    private final CsvOptions options;

    private CsvWriter writer;
    private List<String> cells;

    public CsvSpreadsheetWriter(final OutputStream out) {
        this(out, CsvWriter.GB2312);
    }

    public CsvSpreadsheetWriter(final OutputStream out, final Charset encoding) {
        this.out = out;
        this.options = new CsvOptions(encoding, CsvLineParser.DEFAULT_SEPARATOR, CsvLineParser.DEFAULT_QUOTE_CHARACTER, CsvLineParser.DEFAULT_ESCAPE_CHARACTER);
    }

    @Override
    public SpreadsheetWriter configure(final String option, final Object value) {
        this.options.set(option, value);
        return this;
    }

    @Override
    public SpreadsheetWriter start(final String worksheetName) throws SpreadsheetException {
        if (null != this.writer) {
            throw new SpreadsheetException();
        }
        this.row = 0;
        this.col = 0;
        this.cells = new LinkedList<String>();
        this.writer = options.createWriter(out);
        return this;
    }

    @Override
    public SpreadsheetWriter write(final Boolean bool) throws SpreadsheetException {
        return this.write(null != bool ? bool.toString() : null);
    }

    @Override
    public SpreadsheetWriter write(final Number number) throws SpreadsheetException {
        String text = null;
        if (null != number) {
            final NumberFormat fmt = NumberFormat.getInstance();
            fmt.setGroupingUsed(false);
            text = fmt.format(number, new StringBuffer(), new FieldPosition(0)).toString();
        }
        return this.write(text);
    }

    @Override
    public SpreadsheetWriter write(final Date date) throws SpreadsheetException {
        return write(date, "yyyy-MM-dd HH:mm:ss");
    }

    @Override
    public SpreadsheetWriter write(final Date date, final String pattern) throws SpreadsheetException {
        return this.write(null != date ? new SimpleDateFormat(pattern).format(date) : null);
    }

    @Override
    public SpreadsheetWriter write(final String text) throws SpreadsheetException {
        if (null == writer) {
            throw new SpreadsheetException("not initialized");
        }
        this.cells.add(null != text ? text : "");
        return this;
    }

    @Override
    public AbstractSpreadsheetWriter next() throws SpreadsheetException {
        String[] line = new String[this.cells.size()];
        line = this.cells.toArray(line);
        this.cells.clear();
        this.writer.writeNext(line);
        return super.next();
    }

    @Override
    public void close() throws SpreadsheetException {
        if (!this.cells.isEmpty()) {
            String[] line = new String[this.cells.size()];
            line = this.cells.toArray(line);
            this.cells.clear();
            this.writer.writeNext(line);
        }
        try {
            this.writer.flushQuietly();
            this.writer.close();
        } catch (final IOException e) {
            throw new SpreadsheetException(e);
        }
    }
}
