package com.github.nomou.spreadsheet.msexcel;

import com.github.nomou.spreadsheet.Spreadsheet;
import com.github.nomou.spreadsheet.SpreadsheetException;
import com.github.nomou.spreadsheet.SpreadsheetWriter;
import com.github.nomou.spreadsheet.spi.SpreadsheetWriterFactory;

import java.io.OutputStream;

/**
 * Spreadsheet writer factory for Microsoft Excel 2007+.
 *
 * @author vacoor
 * @since 1.0
 */
public class OpenXMLSpreadsheetWriterFactory implements SpreadsheetWriterFactory {
    private static final boolean POI_OOXML_PRESENT = SpreadsheetImplUtils.isPresent(SpreadsheetImplUtils.SXSSF_STREAMING_CLASS_NAME, SpreadsheetImplUtils.OOXML_CLASS_NAME);

    @Override
    public Spreadsheet.Format[] getSupportedFormats() {
        return new Spreadsheet.Format[]{OpenXMLSpreadsheetParserFactory.OOXML};
    }

    @Override
    public SpreadsheetWriter create(final OutputStream out) throws SpreadsheetException {
        if (POI_OOXML_PRESENT) {
            return new OpenXMLSpreadsheetWriter(out, null);
        }
        throw new SpreadsheetException(String.format("POI '%s' or '%s' missing", SpreadsheetImplUtils.SXSSF_STREAMING_CLASS_NAME, SpreadsheetImplUtils.OOXML_CLASS_NAME));
    }
}
