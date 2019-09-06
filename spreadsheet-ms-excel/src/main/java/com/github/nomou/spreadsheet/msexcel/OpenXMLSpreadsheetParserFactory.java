package com.github.nomou.spreadsheet.msexcel;

import com.github.nomou.spreadsheet.Spreadsheet;
import com.github.nomou.spreadsheet.SpreadsheetException;
import com.github.nomou.spreadsheet.SpreadsheetParser;
import com.github.nomou.spreadsheet.spi.SpreadsheetParserFactory;

import java.io.InputStream;

/**
 * Spreadsheet parser factory for Microsoft Excel 2007+.
 *
 * @author vacoor
 * @since 1.0
 */
public class OpenXMLSpreadsheetParserFactory implements SpreadsheetParserFactory {
    private static final boolean POI_OOXML_PRESENT = SpreadsheetImplUtils.isPresent(SpreadsheetImplUtils.SXSSF_STREAMING_CLASS_NAME, SpreadsheetImplUtils.OOXML_CLASS_NAME);

    /**
     * Microsoft Excel 2007+(*.xlsx) - OPEN-XML.
     * <p>
     * The first 4 bytes of an OOXML file, used in detection
     */
    private static final byte[] OOXML_FILE_HEADER = new byte[]{0x50, 0x4b, 0x03, 0x04};

    public static final Spreadsheet.Format OOXML = new Spreadsheet.Format("Microsoft Excel 2007+", OOXML_FILE_HEADER, "xlsx");

    @Override
    public Spreadsheet.Format[] getSupportedFormats() {
        return new Spreadsheet.Format[]{OOXML};
    }

    @Override
    public SpreadsheetParser create(final InputStream in) throws SpreadsheetException {
        if (POI_OOXML_PRESENT) {
            return new OpenXMLSpreadsheetParser(in);
        }
        throw new SpreadsheetException(String.format("POI '%s' or '%s' missing", SpreadsheetImplUtils.SXSSF_STREAMING_CLASS_NAME, SpreadsheetImplUtils.OOXML_CLASS_NAME));
    }
}
