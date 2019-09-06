package com.github.nomou.spreadsheet.msexcel;

import java.io.*;
import java.lang.Boolean;
import java.lang.Number;
import java.util.Date;

import com.github.nomou.spreadsheet.AbstractSpreadsheetWriter;
import com.github.nomou.spreadsheet.SpreadsheetException;
import com.github.nomou.spreadsheet.SpreadsheetWriter;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;

/**
 * JXL-based spreadsheet writer for BIFF-5(Microsoft Excel 5.0/95) and BIFF-8(Microsoft Excel 97-2003).
 *
 * @author vacoor
 * @see <a href="http://jxl.sourceforge.net/">Java Excel API</a>
 * @since 1.0
 */
class LegacySpreadsheetWriter extends AbstractSpreadsheetWriter {
    public static final String OPTION_TEMPLATE_KEY = "template";

    private OutputStream out;
    private WorkbookSettings settings = new WorkbookSettings();
    private Workbook template;
    private WritableWorkbook workbook;
    private WritableSheet worksheet;

    private int worksheetIndex = -1;

    public LegacySpreadsheetWriter(final OutputStream out) throws SpreadsheetException {
        this(out, null, null);
    }

    public LegacySpreadsheetWriter(final OutputStream out, final InputStream template) throws SpreadsheetException {
        this(out, template, null);
    }

    public LegacySpreadsheetWriter(final OutputStream out, final InputStream template, final String encoding) throws SpreadsheetException {
        this.out = out;

        if (null != encoding) {
            this.settings.setEncoding(encoding);
        }

        if (null != template) {
            try {
                this.template = Workbook.getWorkbook(template, settings);
            } catch (final IOException e) {
                throw new SpreadsheetException(e);
            } catch (final BiffException e) {
                throw new SpreadsheetException(e.getMessage(), e.getCause());
            }
        }
    }

    @Override
    public SpreadsheetWriter configure(final String option, final Object value) {
        if (OPTION_TEMPLATE_KEY.equals(option)) {
            if (value instanceof InputStream) {
                try {
                    template = Workbook.getWorkbook((InputStream) value, settings);
                } catch (final IOException e) {
                    throw new IllegalArgumentException(e);
                } catch (final BiffException e) {
                    throw new IllegalArgumentException(e);
                }
            } else {
                throw new IllegalArgumentException("illegal template value, must be InputStream");
            }
        }
        return this;
    }

    @Override
    public LegacySpreadsheetWriter start(final String worksheetName) throws SpreadsheetException {
        try {
            if (null == workbook) {
                if (null != template) {
                    workbook = Workbook.createWorkbook(out, template, settings);
                } else {
                    workbook = Workbook.createWorkbook(out, settings);
                }
            }

            final WritableSheet[] sheets = workbook.getSheets();
            for (int i = 0; i < sheets.length; i++) {
                final WritableSheet sheet = sheets[i];
                if (sheet.getName().equals(worksheetName)) {
                    worksheet = sheet;
                    worksheetIndex = i;
                    break;
                }
            }

            if (null == worksheet) {
                worksheet = workbook.createSheet(worksheetName, ++worksheetIndex);
            }

            this.row = 0;
            this.col = 0;
            return this;
        } catch (final IOException e) {
            throw new SpreadsheetException(e);
        }
    }

    @Override
    public LegacySpreadsheetWriter write(final Boolean bool) throws SpreadsheetException {
        return null != bool ? doWriteCell(new jxl.write.Boolean(col++, row, bool)) : writeNull(col++, row);
    }

    @Override
    public LegacySpreadsheetWriter write(final Number number) throws SpreadsheetException {
        return null != number ? doWriteCell(new jxl.write.Number(col++, row, number.doubleValue())) : writeNull(col++, row);
    }

    @Override
    public LegacySpreadsheetWriter write(final Date date) throws SpreadsheetException {
        return null != date ? doWriteCell(new jxl.write.DateTime(col++, row, date)) : writeNull(col++, row);
    }

    @Override
    public LegacySpreadsheetWriter write(final Date date, final String pattern) throws SpreadsheetException {
        jxl.write.WritableCellFormat cfmt = new jxl.write.WritableCellFormat(new jxl.write.DateFormat(pattern));
        return null != date ? doWriteCell(new jxl.write.DateTime(col++, row, date, cfmt)) : writeNull(col++, row);
    }

    @Override
    public LegacySpreadsheetWriter write(final String text) throws SpreadsheetException {
        return null != text ? doWriteCell(new jxl.write.Label(col++, row, text)) : writeNull(col++, row);
    }

    public LegacySpreadsheetWriter writeNull(final int col, final int row) throws SpreadsheetException {
        return doWriteCell(new Blank(col, row));
    }

    protected LegacySpreadsheetWriter doWriteCell(final WritableCell cell) throws SpreadsheetException {
        if (null == cell) {
            throw new IllegalArgumentException("cell should non-null");
        }
        if (null == worksheet) {
            throw new IllegalStateException("no writable worksheet at current state, already call start(worksheetName) method?");
        }

        try {
            worksheet.addCell(cell);
            return this;
        } catch (final RowsExceededException e) {
            throw new SpreadsheetException(e.getMessage(), e.getCause());
        } catch (final WriteException e) {
            throw new SpreadsheetException(e.getMessage(), e.getCause());
        }
    }


    @Override
    public void close() throws SpreadsheetException {
        try {
            if (null != workbook) {
                workbook.write();
                workbook.close();
            }
            out.close();
        } catch (WriteException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.gc();
    }
}
