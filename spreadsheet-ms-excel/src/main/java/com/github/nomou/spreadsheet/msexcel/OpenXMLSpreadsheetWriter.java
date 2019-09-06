package com.github.nomou.spreadsheet.msexcel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * POI-based spreadsheet writer for Microsoft Excel 2007+.
 *
 * @author vacoor
 * @see <a href="http://poi.apache.org/">POI</a>
 * @since 1.0
 */
class OpenXMLSpreadsheetWriter extends AbstractPOISpreadsheetWriter {
    private final int rowAccessWindowSize;

    OpenXMLSpreadsheetWriter(final OutputStream out, final InputStream template) {
        this(out, template, 500);
    }

    OpenXMLSpreadsheetWriter(final OutputStream out, final InputStream template, final int rowAccessWindowSize) {
        super(out, template);
        this.rowAccessWindowSize = rowAccessWindowSize;
    }

    @Override
    protected Workbook createWorkbook(final InputStream template) throws IOException {
        if (null != template) {
            try {
                final OPCPackage spreadsheet = OPCPackage.open(template);
                final XSSFWorkbook templateWorkbook = new XSSFWorkbook(spreadsheet);
                return new SXSSFWorkbook(templateWorkbook, this.rowAccessWindowSize);
            } catch (final InvalidFormatException e) {
                throw new IOException(e.getMessage(), e.getCause());
            } catch (final IOException e) {
                throw e;
            }
        } else {
            return new SXSSFWorkbook(this.rowAccessWindowSize);
        }
    }
}
