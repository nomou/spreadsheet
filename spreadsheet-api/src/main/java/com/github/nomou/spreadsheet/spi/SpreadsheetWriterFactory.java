package com.github.nomou.spreadsheet.spi;

import com.github.nomou.spreadsheet.Spreadsheet;
import com.github.nomou.spreadsheet.SpreadsheetException;
import com.github.nomou.spreadsheet.SpreadsheetWriter;

import java.io.OutputStream;

/**
 * Spreadsheet writer factory.
 *
 * @author vacoor
 * @since 1.0
 */
public interface SpreadsheetWriterFactory {

    /**
     * Get the spreadsheet formats supported by the current writer factory.
     *
     * @return the supported spreadsheet formats
     */
    Spreadsheet.Format[] getSupportedFormats();

    /**
     * Create a spreadsheet writer with the given output stream.
     *
     * @param out the output
     * @return the spreadsheet parser
     * @throws SpreadsheetException if there is an error processing the underlying output
     */
    SpreadsheetWriter create(final OutputStream out) throws SpreadsheetException;

}
