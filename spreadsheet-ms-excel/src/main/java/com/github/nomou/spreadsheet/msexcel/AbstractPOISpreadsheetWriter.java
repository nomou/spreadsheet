package com.github.nomou.spreadsheet.msexcel;

import com.github.nomou.spreadsheet.AbstractSpreadsheetWriter;
import com.github.nomou.spreadsheet.SpreadsheetException;
import com.github.nomou.spreadsheet.SpreadsheetWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * Abstract POI spreadsheet writer.
 *
 * @author vacoor
 * @since 1.0
 */
abstract class AbstractPOISpreadsheetWriter extends AbstractSpreadsheetWriter {
    public static final String OPTION_TEMPLATE_KEY = "template";

    /**
     * Template ?
     */
    private InputStream template;

    /**
     * The internal output.
     */
    private OutputStream out;
    private Map<String, CellStyle> fmtStyles = new HashMap<String, CellStyle>();
    private Workbook workbook;
    private Sheet worksheet;
    private int worksheetIndex = -1;

    public AbstractPOISpreadsheetWriter(final OutputStream out, final InputStream template) {
        this.out = out;
        this.template = template;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public SpreadsheetWriter configure(final String option, final Object value) {
        if (OPTION_TEMPLATE_KEY.equals(option)) {
            if (value instanceof InputStream) {
                template = (InputStream) value;
            } else {
                throw new IllegalArgumentException("illegal template value, must be InputStream");
            }
        }
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public AbstractPOISpreadsheetWriter start(final String worksheetName) throws SpreadsheetException {
        try {
            if (null == workbook) {
                workbook = createWorkbook(this.template);
            }

            Sheet sheet = workbook.getSheet(worksheetName);
            if (null == sheet) {
                sheet = workbook.createSheet(worksheetName);
            }

            this.row = 0;
            this.col = 0;
            this.worksheet = sheet;
            this.worksheetIndex = workbook.getSheetIndex(sheet);
        } catch (final IOException e) {
            throw new SpreadsheetException(e);
        }
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public AbstractPOISpreadsheetWriter write(final Boolean bool) throws SpreadsheetException {
        if (null != bool) {
            getCell(col++, row).setCellValue(bool);
        } else {
            writeNull(col++, row);
        }
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public AbstractPOISpreadsheetWriter write(final Number number) throws SpreadsheetException {
        if (null != number) {
            getCell(col++, row).setCellValue(number.doubleValue());
        } else {
            writeNull(col++, row);
        }
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public AbstractPOISpreadsheetWriter write(final Date date) throws SpreadsheetException {
        return write(date, "yyyy-MM-dd HH:mm:ss");
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public AbstractPOISpreadsheetWriter write(final Date date, final String pattern) throws SpreadsheetException {
        if (null != date) {
            final Cell cell = getCell(col++, row);
            cell.setCellValue(date);
            cell.setCellStyle(getDateStyle(pattern));
        } else {
            writeNull(col++, row);
        }
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public AbstractPOISpreadsheetWriter write(final String text) throws SpreadsheetException {
        getCell(col++, row).setCellValue(null != text ? text : "");
        return this;
    }

    public AbstractPOISpreadsheetWriter writeNull(final int col, final int row) throws SpreadsheetException {
        getCell(col, row).setCellValue("");
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public AbstractPOISpreadsheetWriter next() {
        row++;
        col = 0;
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void close() {
        try {
            if (null != workbook && null != out) {
                workbook.write(out);
                System.gc();
            }
        } catch (final IOException e) {
            throw new RuntimeException(e);
        }
    }

    protected Cell getCell(final int col, final int row) {
        if (null == worksheet) {
            throw new IllegalStateException("no writable worksheet at current state, already call start(worksheetName) method?");
        }
        Row r = worksheet.getRow(row);
        if (null == r) {
            r = worksheet.createRow(row);
        }
        Cell c = r.getCell(col);
        if (null == c) {
            c = r.createCell(col);
        }
        return c;
    }

    protected CellStyle getDateStyle(final String format) {
        if (null == worksheet) {
            throw new IllegalStateException("no writable worksheet at current state, already call start(worksheetName) method?");
        }

        CellStyle cellStyle = fmtStyles.get(format);
        if (null == cellStyle) {
            final Workbook workbook = worksheet.getWorkbook();
            cellStyle = workbook.createCellStyle();
            cellStyle.setDataFormat(workbook.createDataFormat().getFormat(format));
            fmtStyles.put(format, cellStyle);
        }
        return cellStyle;
    }

    protected abstract Workbook createWorkbook(final InputStream template) throws IOException;
}
