package com.github.nomou.spreadsheet.msexcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * POI-based spreadsheet writer for Microsoft Excel 97-2003.
 * <p>
 * WARN: Not support BIFF5 (Microsoft Excel 5.0/95).
 * </p>
 *
 * @author vacoor
 * @see <a href="http://poi.apache.org/">POI</a>
 * @since 1.0
 */
class LegacySpreadsheetWriter2 extends AbstractPOISpreadsheetWriter {

    public LegacySpreadsheetWriter2(final OutputStream out, final InputStream template) {
        super(out, template);
    }

    @Override
    protected Workbook createWorkbook(final InputStream template) throws IOException {
        return null != template ? new HSSFWorkbook(template) : new HSSFWorkbook();
    }
}
