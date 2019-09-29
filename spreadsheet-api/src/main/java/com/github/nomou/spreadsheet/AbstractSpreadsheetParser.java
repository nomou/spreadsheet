package com.github.nomou.spreadsheet;

import java.util.ArrayList;
import java.util.List;
import java.util.NoSuchElementException;

/**
 * Abstract spreadsheet parser.
 *
 * @author vacoor
 * @since 1.0
 */
public abstract class AbstractSpreadsheetParser implements SpreadsheetParser {
    /**
     * The current parsing event type.
     */
    protected int eventType;

    /**
     * The current worksheet index.
     */
    protected int worksheetIndex = -1;

    /**
     * The current worksheet name.
     */
    protected String worksheetName;

    /**
     * {@inheritDoc}
     */
    @Override
    public int getEventType() {
        return this.eventType;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getWorksheetIndex() {
        final int st = this.eventType;
        if (START_WORKBOOK == st || END_WORKBOOK == st) {
            throw new IllegalStateException("getWorksheetIndex() called in illegal state");
        }
        return this.worksheetIndex;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getWorksheetName() {
        final int st = this.eventType;
        if (START_WORKBOOK == st || END_WORKBOOK == st) {
            throw new IllegalStateException("getWorksheetName() called in illegal state");
        }
        return this.worksheetName;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean hasNext() {
        // returns -1 when it detects a broken stream
        return EOF != this.eventType && END_WORKBOOK != this.eventType;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int next() throws SpreadsheetException {
        final int t = this.eventType;
        if (!hasNext()) {
            if (EOF != t) {
                throw new NoSuchElementException("END_WORKBOOK reached: no more elements on the stream.");
            } else {
                throw new IllegalStateException("Error processing input source. The input stream is not complete.");
            }
        }
        return this.eventType = doNext();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Object[] nextRecord(final boolean ignoreEmptyRecord) throws SpreadsheetException {
        final SpreadsheetParser parser = this;

        final List<Object> values = new ArrayList<Object>();
        boolean noMoreRecord = true;
        while (parser.hasNext()) {
            final int event = parser.next();
            if (END_CELL == event) {
                noMoreRecord = false;
                final int col = parser.getCol();
                final Object value = parser.getValue();

                if (col < values.size() - 1) {
                    throw new IllegalStateException("illegal column: " + col);
                }

                if (null != value) {
                    fillNull(values, col - values.size());
                    values.add(value);
                }
            } else if (END_RECORD == event) {
                if (!ignoreEmptyRecord || !values.isEmpty()) {
                    noMoreRecord = false;
                    break;
                }
                noMoreRecord = true;
                values.clear();
            }
        }

        return noMoreRecord ? null : values.toArray(new Object[values.size()]);
    }

    private void fillNull(final List<?> values, final int count) {
        for (int i = 0; i < count; i++) {
            values.add(null);
        }
    }

    /**
     * Parse next event.
     *
     * @return next event type
     * @throws SpreadsheetException if parsing error
     */
    protected abstract int doNext() throws SpreadsheetException;
}
