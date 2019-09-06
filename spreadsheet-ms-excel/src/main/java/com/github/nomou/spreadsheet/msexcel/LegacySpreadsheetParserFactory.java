package com.github.nomou.spreadsheet.msexcel;

import com.github.nomou.spreadsheet.Spreadsheet;
import com.github.nomou.spreadsheet.SpreadsheetException;
import com.github.nomou.spreadsheet.SpreadsheetParser;
import com.github.nomou.spreadsheet.spi.SpreadsheetParserFactory;

import java.io.InputStream;

/**
 * Spreadsheet parser factory for Microsoft Excel 95-2003.
 *
 * @author vacoor
 * @since 1.0
 */
public class LegacySpreadsheetParserFactory implements SpreadsheetParserFactory {
    private static final boolean JXL_PRESENT = SpreadsheetImplUtils.isPresent(SpreadsheetImplUtils.JXL_CLASS_NAME);
    private static final boolean POI_HSSF_PRESENT = SpreadsheetImplUtils.isPresent(SpreadsheetImplUtils.HSSF_STREAMING_CLASS_NAME);

    /**
     * Microsoft Excel 5.0/95, Microsoft Excel 97-2003(*.xls).
     * <p>
     * OLE2 / BIFF8+ stream used for Office 97 and higher documents.
     */
    private static final byte[] OLE2_FILE_HEADER = new byte[]{(byte) 0xD0, (byte) 0xCF, 0x11, (byte) 0xE0, (byte) 0xA1, (byte) 0xB1, 26, (byte) 0xE1};

    public static final Spreadsheet.Format OLE2 = new Spreadsheet.Format("Microsoft Excel 5.0/95, 97-2003", OLE2_FILE_HEADER, "xls");

    @Override
    public Spreadsheet.Format[] getSupportedFormats() {
        return new Spreadsheet.Format[]{OLE2};
    }

    @Override
    public SpreadsheetParser create(final InputStream in) throws SpreadsheetException {
        if (JXL_PRESENT) {
            return new LegacySpreadsheetParser(in);
        }
        if (POI_HSSF_PRESENT) {
            return new LegacySpreadsheetParser2(in);
        }
        throw new SpreadsheetException("jxl and POI HSSF missing");
    }
}
