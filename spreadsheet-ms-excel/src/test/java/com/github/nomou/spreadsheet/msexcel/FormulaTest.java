package com.github.nomou.spreadsheet.msexcel;

import com.github.nomou.spreadsheet.SpreadsheetParser;
import com.github.nomou.spreadsheet.util.SpreadsheetRecord;

import java.io.InputStream;
import java.util.Arrays;

/**
 * Created by vacoor on 2019/9/29.
 */
public class FormulaTest {
    public static void main(String[] args) {
        final InputStream in = FormulaTest.class.getResourceAsStream("/formula.xls");
        final SpreadsheetParser parser = new LegacySpreadsheetParser2(in);
//        final InputStream in = FormulaTest.class.getResourceAsStream("/formula.xlsx");
//        final SpreadsheetParser parser = new OpenXMLSpreadsheetParser(in);
        try {
            do {
                final Object[] record = parser.nextRecord(true);
                if (null != record) {
                    final int row = parser.getRow();
                    final int sheet = parser.getWorksheetIndex();

                    System.out.println("sheet" + sheet + ":" + row + ":" + Arrays.toString(record));
                }
            } while (SpreadsheetParser.END_WORKBOOK != parser.getEventType());
        } finally {
            parser.close();
        }
    }
}
