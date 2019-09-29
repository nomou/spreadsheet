package com.github.nomou.spreadsheet.msexcel;

import com.github.nomou.spreadsheet.AbstractSpreadsheetParser;
import com.github.nomou.spreadsheet.SpreadsheetException;
import com.github.nomou.spreadsheet.SpreadsheetParser;
import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingRowDummyRecord;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Queue;

/**
 * POI-based spreadsheet parser for Microsoft Excel 97-2003.
 * <p>
 * WARN: Not support BIFF5 (Microsoft Excel 5.0/95).
 * </p>
 *
 * @author vacoor
 * @see <a href="http://poi.apache.org/">POI</a>
 * @since 1.0
 */
class LegacySpreadsheetParser2 extends AbstractSpreadsheetParser {
    private WorkbookIterator workbookIt;
    private List<BoundSheetRecord> boundSheetRecords;
    private SSTRecord sharedStyleTable;

    private int row = -1;
    private int col = -1;
    private Object value;
    private Record _next;

    LegacySpreadsheetParser2(final InputStream inputStream) throws SpreadsheetException {
        setInputSource(inputStream);
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
        return null != this.boundSheetRecords ? this.boundSheetRecords.size() : 0;
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


    @Override
    protected int doNext() throws SpreadsheetException {
        final int t = this.eventType;
        final WorkbookIterator it = this.workbookIt;
        final int worksheets = getNumberOfWorksheets();
        final int worksheetIndex = this.worksheetIndex;

        // FIX START_RECORD.
        if (START_RECORD == t && null != this._next) {
            if (this._next instanceof CellValueRecordInterface) {
                final CellValueRecordInterface cell = (CellValueRecordInterface) this._next;
                this.col = cell.getColumn();
                this.row = cell.getRow();
                this.value = asJavaObject(cell, it);
            } else if (this._next instanceof MissingCellDummyRecord) {
                final MissingCellDummyRecord cell = (MissingCellDummyRecord) this._next;
                this.col = cell.getColumn();
                this.row = cell.getRow();
                this.value = null;
            }
            this._next = null;
            return START_CELL;
        }

        // FIX END_CELL.
        if (START_CELL == t) {
            return END_CELL;
        }

        // FIX END_WORKBOOK
        if (END_WORKSHEET == t && !it.hasNext()) {
            doPostWorkbook();
            return END_WORKBOOK;
        }

        int newEvent = EOF;
        while (it.hasNext()) {
            final Record record = it.next();
            /*-
             * --- workbook BOF record
             * --- ...
             * --- worksheets BOF record
             * --- workbook shared style table record
             * --- workbook EOF
             * --- worksheet record        <----
             * --- row record
             * --- worksheet EOF
             * --- EOF
             */
            if (record instanceof EOFRecord) {
                if (START_WORKBOOK == t) {
                    newEvent = EOF;
                } else if (END_WORKSHEET == t && (worksheets - 1 == worksheetIndex)) {
                    doPostWorkbook();
                    newEvent = END_WORKBOOK;
                } else if (END_WORKSHEET != t) {
                    newEvent = END_WORKSHEET;
                }
            } else if ((record instanceof BOFRecord) && isWorkbook((BOFRecord) record)) {
                // workbook start.
                newEvent = EOF;
            } else if (record instanceof BoundSheetRecord) {
                // bound sheet.
                newEvent = EOF;
            } else if (record instanceof SSTRecord) {
                // shared style recod.
                newEvent = EOF;
            } else if ((record instanceof BOFRecord) && isWorksheet((BOFRecord) record)) {
                // worksheet start.
                this.worksheetIndex++;
                this.worksheetName = this.boundSheetRecords.get(this.worksheetIndex).getSheetname();
                newEvent = START_WORKSHEET;
            } else if (record instanceof RowRecord) {
                // FIXME.
                // this.row = ((RowRecord) record).getRowNumber();
                // newEvent = START_RECORD;
//                System.out.println("-----------------");
//                newEvent = EOF;
            } else if (record instanceof CellValueRecordInterface) {
                final CellValueRecordInterface cell = (CellValueRecordInterface) record;
                final int row = cell.getRow();
                final short column = cell.getColumn();
                if (0 == column) {
                    this.row = row;
                    this._next = record;
                    newEvent = START_RECORD;
                } else {
                    this.col = cell.getColumn();
                    this.value = asJavaObject(cell, it);
//                System.out.println("--");
                    // } else if (record instanceof CellRecord) {
                    newEvent = START_CELL;
                }
            } else if (record instanceof MissingRowDummyRecord) {
                // this.row = ((MissingRowDummyRecord) record).getRowNumber();

                // newEvent = START_RECORD;
//                newEvent = EOF;
            } else if (record instanceof MissingCellDummyRecord) {
                final MissingCellDummyRecord cell = (MissingCellDummyRecord) record;
                final int row = cell.getRow();
                final int column = cell.getColumn();

                if (0 == column) {
                    this.row = row;
                    this._next = cell;
                    newEvent = START_RECORD;
                } else {
                    this.col = cell.getColumn();
                    this.value = null;
                    newEvent = START_CELL;
                }
            } else if (record instanceof LastCellOfRowDummyRecord) {
                // TODO
                newEvent = END_RECORD;
            } else {
                newEvent = EOF;
                // TODO
                // ignore
            }

            if (EOF != newEvent) {
                break;
            }
        }
        return newEvent;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void close() throws SpreadsheetException {
        if (END_WORKBOOK != this.eventType) {
            this.eventType = EOF;
        }
        doPostWorkbook();
    }

    void setInputSource(final InputStream inputSource) throws SpreadsheetException {
        try {
            /*-
             * reference: {@link org.apache.poi.hssf.eventusermodel.HSSFEventFactory}
             */
            final POIFSFileSystem fs = new POIFSFileSystem(inputSource);
            final InputStream docIn = fs.getRoot().createDocumentInputStream("Workbook");

            this.workbookIt = new WorkbookIterator(new RecordFactoryInputStream(docIn, false));
            this.doPreWorkbook();
            this.eventType = START_WORKBOOK;
        } catch (final IOException e) {
            throw new SpreadsheetException(e);
        }
    }

    void doPreWorkbook() {
        final WorkbookIterator it = this.workbookIt;
        final List<BoundSheetRecord> boundSheetRecords = new ArrayList<BoundSheetRecord>();
        SSTRecord sharedStyleTable = null;
        while (it.hasNext()) {
            final Record record = it.next();
            if ((record instanceof BOFRecord) && isWorkbook((BOFRecord) record)) {
                // workbook start.
            } else if (record instanceof BoundSheetRecord) {
                boundSheetRecords.add((BoundSheetRecord) record);
            } else if (record instanceof SSTRecord) {
                sharedStyleTable = (SSTRecord) record;
            } else if (record instanceof EOFRecord) {
                break;
            }
        }
        this.boundSheetRecords = Collections.unmodifiableList(boundSheetRecords);
        this.sharedStyleTable = sharedStyleTable;
    }

    void doPostWorkbook() {
        this.worksheetIndex = -1;
        this.worksheetName = null;

        this.sharedStyleTable = null;
        this.boundSheetRecords = null;
        this.workbookIt = null;
    }

    private Object asJavaObject(final CellValueRecordInterface cell, final WorkbookIterator it) {
        final WorkbookIterator workbook = this.workbookIt;

        Object ret = null;
        if (null == cell) {
            ret = null;
        } else if (cell instanceof BoolErrRecord) {
            ret = ((BoolErrRecord) cell).getBooleanValue();
        } else if (cell instanceof NumberRecord) {
            final NumberRecord numberic = (NumberRecord) cell;
            final double value = numberic.getValue();
            ret = workbook.isDateRecord(numberic) ? HSSFDateUtil.getJavaDate(value) : value;
        } else if (cell instanceof LabelRecord) {
            ret = ((LabelRecord) cell).getValue();
            // ret = null != ret ? ret.trim() : null;    // 这里 trim 下, 兼容一下多个换行转换为其他类型出错问题
        } else if (cell instanceof LabelSSTRecord) { // 引用共享字符串表的 label 类型
            final LabelSSTRecord labelSST = (LabelSSTRecord) cell;
            ret = sharedStyleTable.getString(labelSST.getSSTIndex()).getString();
            // ret = null != ret ? ret.trim() : null;    // 这里 trim 下, 兼容一下多个换行转换为其他类型出错问题
        } else if (cell instanceof FormulaRecord) {
            final FormulaRecord formula = (FormulaRecord) cell;
            final int resultType = formula.getCachedResultType();

            // 如果公式的值是一个字符串, 则结果存在下一个 record
            if (HSSFCell.CELL_TYPE_BLANK == resultType || HSSFCell.CELL_TYPE_ERROR == resultType) {
                ret = null;
            } else if (HSSFCell.CELL_TYPE_BOOLEAN == resultType) {
                formula.getCachedBooleanValue();
                ret = 0 != Double.valueOf(formula.getValue()).intValue();
            } else if (HSSFCell.CELL_TYPE_STRING == resultType) {
                /*-
                 * true if this FormulaRecord is followed by a StringRecord representing the cached text result of the formula evaluation.
                 */
                if (!formula.hasCachedResultString()) {
                    throw new IllegalStateException("text formula not has cached result string");
                } else if (!it.hasNext()) {
                    throw new IllegalStateException("text formula record is not flowed string record");
                } else {
                    final Record record = it.next();
                    if (!(record instanceof StringRecord)) {
                        throw new IllegalStateException("text formula record is not flowed string record");
                    }
                    ret = ((StringRecord) record).getString();
                }
            } else {    // 默认为 Number
                final double value = formula.getValue();
                ret = workbook.isDateRecord(formula) ? HSSFDateUtil.getJavaDate(value) : value;
            }
        }

        return ret;
    }

    private static boolean isWorkbook(final BOFRecord record) {
        return BOFRecord.TYPE_WORKBOOK == record.getType();
    }

    private static boolean isWorksheet(final BOFRecord record) {
        return BOFRecord.TYPE_WORKSHEET == record.getType();
    }

    private static class WorkbookIterator implements Iterator<Record> {
        private final RecordFactoryInputStream recordFactory;
        private final Queue<Record> out = new ArrayDeque<Record>();
        private FormatTrackingHSSFListener delegate = new FormatTrackingHSSFListener(new MissingRecordAwareHSSFListener(new HSSFListener() {
            @Override
            public void processRecord(final Record r) {
                out.offer(r);
            }
        }));

        WorkbookIterator(final RecordFactoryInputStream recordFactory) {
            this.recordFactory = recordFactory;
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public boolean hasNext() {
            if (!out.isEmpty()) {
                return true;
            }

            final Record record = recordFactory.nextRecord();
            if (null != record) {
                this.delegate.processRecord(record);
            }
            return !out.isEmpty();
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public Record next() {
            if (!hasNext()) {
                throw new NoSuchElementException("no more elements on the stream.");
            }
            return this.out.poll();
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public void remove() {
            throw new UnsupportedOperationException("Method not supported!");
        }

        /**
         * Returns true if given cell is a date record.
         *
         * @param cell cell.
         * @return true if the cell is date record.
         */
        boolean isDateRecord(final CellValueRecordInterface cell) {
            final int formatIndex = this.delegate.getFormatIndex(cell);
            final String formatString = this.delegate.getFormatString(cell);
            return HSSFDateUtil.isADateFormat(formatIndex, formatString);
        }
    }
}
