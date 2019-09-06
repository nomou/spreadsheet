package com.github.nomou.spreadsheet.csv;

import com.github.nomou.spreadsheet.Spreadsheet;
import com.github.nomou.spreadsheet.SpreadsheetParser;
import com.github.nomou.spreadsheet.spi.SpreadsheetParserFactory;

import java.io.InputStream;

/**
 * CSV spreadsheet parser factory.
 *
 * @author vacoor
 * @since 1.0
 */
public class CsvSpreadsheetParserFactory implements SpreadsheetParserFactory {
    public static final Spreadsheet.Format CSV = new Spreadsheet.Format("CSV", new byte[0], "csv");

    @Override
    public Spreadsheet.Format[] getSupportedFormats() {
        return new Spreadsheet.Format[]{CSV};
    }

    @Override
    public SpreadsheetParser create(final InputStream in) {
        return new CsvSpreadsheetParser(in);
    }
}
