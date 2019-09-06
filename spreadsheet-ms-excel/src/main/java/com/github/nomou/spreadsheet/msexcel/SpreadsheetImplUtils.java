package com.github.nomou.spreadsheet.msexcel;

/**
 * impl util.
 *
 * @author vacoor
 * @since 1.0
 */
class SpreadsheetImplUtils {
    static final String JXL_CLASS_NAME = "jxl.Workbook";

    static final String HSSF_STREAMING_CLASS_NAME = "org.apache.poi.hssf.record.RecordFactoryInputStream";

    static final String SXSSF_STREAMING_CLASS_NAME = "org.apache.poi.xssf.streaming.SXSSFWorkbook";

    static final String OOXML_CLASS_NAME = "org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook";

    private SpreadsheetImplUtils() {
    }

    private static boolean isPresent(final String className) {
        try {
            Class.forName(className);
            return true;
        } catch (final ClassNotFoundException e) {
            return false;
        }
    }

    static boolean isPresent(final String... classNames) {
        for (final String className : classNames) {
            if (!isPresent(className)) {
                return false;
            }
        }
        return true;
    }
}
