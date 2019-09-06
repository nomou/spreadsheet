package com.github.nomou.spreadsheet.csv;

import com.github.nomou.spreadsheet.Spreadsheet;
import com.github.nomou.spreadsheet.SpreadsheetWriter;
import com.github.nomou.spreadsheet.spi.SpreadsheetWriterFactory;

import java.io.OutputStream;

/**
 * CSV spreadsheet writer factory.
 *
 * @author vacoor
 * @since 1.0
 */
public class CsvSpreadsheetWriterFactory implements SpreadsheetWriterFactory {

    @Override
    public Spreadsheet.Format[] getSupportedFormats() {
        return new Spreadsheet.Format[]{CsvSpreadsheetParserFactory.CSV};
    }

    @Override
    public SpreadsheetWriter create(final OutputStream out) {
        return new CsvSpreadsheetWriter(out);
    }
}
