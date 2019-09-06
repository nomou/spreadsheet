package com.github.nomou.spreadsheet.msexcel;

import com.github.nomou.spreadsheet.Spreadsheet;
import com.github.nomou.spreadsheet.SpreadsheetException;
import com.github.nomou.spreadsheet.SpreadsheetWriter;
import com.github.nomou.spreadsheet.spi.SpreadsheetWriterFactory;

import java.io.OutputStream;

/**
 * Spreadsheet writer factory for Microsoft Excel 95-2003.
 *
 * @author vacoor
 * @since 1.0
 */
public class LegacySpreadsheetWriterFactory implements SpreadsheetWriterFactory {
    private static final boolean JXL_PRESENT = SpreadsheetImplUtils.isPresent(SpreadsheetImplUtils.JXL_CLASS_NAME);
    private static final boolean POI_HSSF_PRESENT = SpreadsheetImplUtils.isPresent(SpreadsheetImplUtils.HSSF_STREAMING_CLASS_NAME);

    @Override
    public Spreadsheet.Format[] getSupportedFormats() {
        return new Spreadsheet.Format[]{LegacySpreadsheetParserFactory.OLE2};
    }

    @Override
    public SpreadsheetWriter create(final OutputStream out) throws SpreadsheetException {
        if (JXL_PRESENT) {
            return new LegacySpreadsheetWriter(out);
        }
        if (POI_HSSF_PRESENT) {
            return new LegacySpreadsheetWriter2(out, null);
        }
        throw new IllegalStateException("jxl and POI HSSF missing");
    }
}
