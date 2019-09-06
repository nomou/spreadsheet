package com.github.nomou.spreadsheet.msexcel;

import com.github.nomou.spreadsheet.AbstractSpreadsheetParser;
import com.github.nomou.spreadsheet.SpreadsheetException;
import com.github.nomou.spreadsheet.SpreadsheetParser;
import jxl.BooleanCell;
import jxl.Cell;
import jxl.CellType;
import jxl.DateCell;
import jxl.LabelCell;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.NoSuchElementException;
import java.util.TimeZone;

/**
 * JXL-based spreadsheet parser for BIFF-5(Microsoft Excel 5.0/95) and BIFF-8(Microsoft Excel 97-2003).
 *
 * @author vacoor
 * @see <a href="http://jxl.sourceforge.net/">Java Excel API</a>
 * @since 1.0
 */
class LegacySpreadsheetParser extends AbstractSpreadsheetParser {
    private int row = -1;
    private int col = -1;
    private Object value;

    private Workbook workbook;

    LegacySpreadsheetParser(final InputStream in) throws SpreadsheetException {
        initInputSource(in);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public SpreadsheetParser configure(final String option, final Object value) {
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getNumberOfWorksheets() {
        return this.workbook.getNumberOfSheets();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getRow() {
        final int st = eventType;
        if (START_RECORD != st && END_RECORD != st && START_CELL != st && END_CELL != st) {
            throw new IllegalStateException("getRow() called in illegal state");
        }
        return this.row;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getCol() {
        final int st = this.eventType;
        if (START_CELL != st && END_CELL != st) {
            throw new IllegalStateException("getCol() called in illegal state");
        }
        return this.col;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Object getValue() {
        final int st = this.eventType;
        if (START_CELL != st && END_CELL != st) {
            throw new IllegalStateException("getValue() called in illegal state");
        }
        return this.value;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected int doNext() throws SpreadsheetException {
        final Workbook workbook = this.workbook;
        final int worksheets = workbook.getNumberOfSheets();
        final int worksheetIndex = this.worksheetIndex;
        final int st = this.eventType;

        int newEvent = EOF;
        if (START_WORKBOOK == st) {
            this.worksheetIndex = 0;
            if (1 > worksheets) {
//                close();
                newEvent = END_WORKBOOK;
            } else {
                newEvent = START_WORKSHEET;
            }
        } else if (START_WORKSHEET == st) {
            final Sheet sheet = workbook.getSheet(worksheetIndex);
            final int rows = sheet.getRows();
            if (1 > rows) {
                newEvent = END_WORKSHEET;
            } else {
                row = -1;
                Cell[] cells;
                do {
                    cells = sheet.getRow(++row);
                } while (null != cells && 1 > cells.length);

                newEvent = START_RECORD;
            }
        } else if (START_RECORD == st) {
            col = 0;
            final Sheet sheet = workbook.getSheet(worksheetIndex);
            Cell[] cells = sheet.getRow(this.row);
            if (col >= cells.length) {
                throw new NoSuchElementException();
            }
            value = asJavaObject(cells[col]);

            newEvent = START_CELL;
        } else if (START_CELL == st) {
            newEvent = END_CELL;
        } else if (END_CELL == st) {
            final Sheet sheet = workbook.getSheet(worksheetIndex);
            Cell[] cells = sheet.getRow(this.row);
            col++;
            // if (col == cells.length - 1) {
            if (col == cells.length) {
                col = -1;
                newEvent = END_RECORD;
            } else {
                if (col >= cells.length) {
                    throw new NoSuchElementException();
                }
                value = asJavaObject(cells[col]);
                newEvent = START_CELL;
            }
        } else if (END_RECORD == st) {
            final int rows = workbook.getSheet(worksheetIndex).getRows();
            if (this.row == rows - 1) {
//                close();
                newEvent = END_WORKSHEET;
            } else {
                row++;
                newEvent = START_RECORD;
            }
        } else if (END_WORKSHEET == st && worksheetIndex == worksheets - 1) {
//            close();
            newEvent = END_WORKBOOK;
        }
        return newEvent;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void close() throws SpreadsheetException {
        if (null != this.workbook) {
            this.workbook.close();
        }
        if (END_WORKBOOK != this.eventType) {
            this.eventType = EOF;
        }
        this.col = -1;
        this.row = -1;
        this.workbook = null;
    }

    /**
     * Configure input source for the parser.
     *
     * @param inputSource input source.
     * @throws IllegalArgumentException if this input source not support
     * @throws SpreadsheetException     If this input source is parsed incorrectly
     */
    public void initInputSource(final Object inputSource) throws SpreadsheetException {
        try {
            final WorkbookSettings settings = new WorkbookSettings();
            settings.setEncoding("CP936");  // TODO biff5 must.

            Workbook workbook;
            if (inputSource instanceof InputStream) {
                workbook = Workbook.getWorkbook((InputStream) inputSource, settings);
            } else if (inputSource instanceof File) {
                workbook = Workbook.getWorkbook((File) inputSource, settings);
            } else {
                throw new IllegalArgumentException("Unsupported input source: " + inputSource);
            }

            this.workbook = workbook;
            this.eventType = START_WORKBOOK;
        } catch (final BiffException e) {
            throw new SpreadsheetException(e.getMessage(), e.getCause());
        } catch (final IOException e) {
            throw new SpreadsheetException(e);
        }
    }

    /**
     * Returns representation of the value of Cell in the Java.
     *
     * @param cell cell.
     * @return representation value in the Java
     */
    private Object asJavaObject(final Cell cell) {
        final CellType type = cell.getType();

        Object value = null;
        if (CellType.EMPTY == type) {
            value = null;
        } else if (CellType.BOOLEAN_FORMULA == type || CellType.BOOLEAN == type) {
            final BooleanCell bool = (BooleanCell) cell;
            value = bool.getValue();
        } else if (CellType.LABEL == type || CellType.STRING_FORMULA == type) {
            final LabelCell label = (LabelCell) cell;
            value = label.getString();
            // value = null != value ? value.trim() : null;    // 这里 trim 下, 兼容一下多个换行转换为其他类型出错问题
        } else if (type == CellType.NUMBER || type == CellType.NUMBER_FORMULA) {
            final NumberCell number = (NumberCell) cell;
            value = number.getValue();
        } else if (type == CellType.DATE || type == CellType.DATE_FORMULA) {
            final DateCell date = (DateCell) cell;
            // 时区为 GMT, Excel 使用GMT时区, Java 会使用默认时区, 这里不做处理了, 直接不允许使用时间
            final Calendar calendar = new GregorianCalendar(TimeZone.getTimeZone("GMT"));
            calendar.setTime(date.getDate());
            calendar.add(Calendar.MILLISECOND, -TimeZone.getDefault().getRawOffset());
            value = calendar.getTime();
        /*
        } else if (type == CellType.ERROR || type == CellType.FORMULA_ERROR) {
            // ErrorFormulaCell error = (ErrorFormulaCell) cell;
            ErrorCell error = (ErrorCell) cell;
            System.out.println(error.getContents());
        }
        */
        } else {
            value = cell.getContents();
        }

        return value;
    }
}
