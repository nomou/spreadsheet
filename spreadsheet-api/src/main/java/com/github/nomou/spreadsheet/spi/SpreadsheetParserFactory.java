package com.github.nomou.spreadsheet.spi;

import com.github.nomou.spreadsheet.Spreadsheet;
import com.github.nomou.spreadsheet.SpreadsheetException;
import com.github.nomou.spreadsheet.SpreadsheetParser;

import java.io.InputStream;

/**
 * Spreadsheet parser factory.
 *
 * @author vacoor
 * @since 1.0
 */
public interface SpreadsheetParserFactory {

    /**
     * Get the spreadsheet formats supported by the current parser factory.
     *
     * @return the supported spreadsheet formats
     */
    Spreadsheet.Format[] getSupportedFormats();

    /**
     * Create a spreadsheet parser with the given input stream.
     *
     * @param in the input source
     * @return the spreadsheet parser
     * @throws SpreadsheetException if there is an error processing the underlying input source
     */
    SpreadsheetParser create(final InputStream in) throws SpreadsheetException;

}
