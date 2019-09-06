package com.github.nomou.spreadsheet;

import java.lang.reflect.Array;
import java.util.Calendar;
import java.util.Date;

/**
 * Abstract spreadsheet writer.
 *
 * @author vacoor
 * @since 1.0
 */
public abstract class AbstractSpreadsheetWriter implements SpreadsheetWriter {
    /**
     * The current row.
     */
    protected int row;

    /**
     * The current column.
     */
    protected int col;

    /**
     * The unnamed sheet counter.
     */
    protected int unnamedCount = 1;

    /**
     * {@inheritDoc}
     */
    @Override
    public SpreadsheetWriter start() throws SpreadsheetException {
        return this.start("sheet" + (unnamedCount++));
    }

    /**
     * {@inheritDoc}
     */
    @Override
    @SuppressWarnings("unckecked")
    public <E> SpreadsheetWriter write(final E... cells) throws SpreadsheetException {
        return write((Object) cells);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public SpreadsheetWriter write(final Object obj) throws SpreadsheetException {
        if (null == obj) {
            this.write("");
            return this;
        }

        if (obj.getClass().isArray()) {
            final int length = Array.getLength(obj);

            for (int i = 0; i < length; i++) {
                this.write(Array.get(obj, i));
            }
            return this;
        }

        if (obj instanceof Boolean) {
            write((Boolean) obj);
        } else if (obj instanceof Number) {
            write((Number) obj);
        } else if (obj instanceof Date) {
            write((Date) obj);
        } else if (obj instanceof Calendar) {
            write(((Calendar) obj).getTime());
        } else {
            write(String.valueOf(obj));
        }
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public AbstractSpreadsheetWriter next() throws SpreadsheetException {
        this.row++;
        this.col = 0;
        return this;
    }
}
