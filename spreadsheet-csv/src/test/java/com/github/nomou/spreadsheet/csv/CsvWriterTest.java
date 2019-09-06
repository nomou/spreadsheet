package com.github.nomou.spreadsheet.csv;

import com.github.nomou.spreadsheet.SpreadsheetException;
import com.github.nomou.spreadsheet.SpreadsheetParser;
import com.github.nomou.spreadsheet.SpreadsheetWriter;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;

/**
 */
public class CsvWriterTest {
    @Test
    public void testWrite() throws IOException, SpreadsheetException {
        SpreadsheetWriter writer = new CsvSpreadsheetWriter(new FileOutputStream("out.csv")).configure("encoding", "UTF-8");
        writer.start().write("Name", "Age").next().write("张三", "18").close();
    }

    @Test
    public void testRead() throws IOException, SpreadsheetException {
        final SpreadsheetParser parser = new CsvSpreadsheetParser(new FileInputStream("out.csv"), CsvWriter.GB2312);
        while (parser.hasNext()) {
            final Object[] row = parser.nextRecord(true);
            System.out.println(Arrays.toString(row));
        }
        System.out.println("-----------");
    }

    @Test
    public void testRead2() throws FileNotFoundException, SpreadsheetException {
        final FileInputStream in = new FileInputStream("out.csv");
        final CsvSpreadsheetParser parser = new CsvSpreadsheetParser(in);
        Object[] values;
        while (null != (values = parser.nextRecord(false))) {
            System.out.println(Arrays.toString(values));
            System.out.println("=================");
        }
//        doConsume(parser);
    }

    @Test
    public void testWrite2() throws FileNotFoundException, SpreadsheetException {
        final FileOutputStream out = new FileOutputStream("out.csv");
        final CsvSpreadsheetWriter writer = new CsvSpreadsheetWriter(out);
        writer.start()
                .write("name", "age").next()
                .write("张三", "18");
        writer.close();
    }


    static void doConsume(final SpreadsheetParser parser) throws SpreadsheetException {
        int eventType;
        do {
            eventType = parser.getEventType();
            if (SpreadsheetParser.START_WORKBOOK == eventType) {
                System.out.println("开始解析 WORKBOOK, 发现 Worksheet: " + parser.getNumberOfWorksheets());
            } else if (SpreadsheetParser.START_WORKSHEET == eventType) {
                System.out.println("开始解析 Worksheet: " + parser.getWorksheetIndex() + ", " + parser.getWorksheetName());
            } else if (SpreadsheetParser.START_RECORD == eventType) {
                // System.out.println("开始解析行:" + parser.getRow());
                System.out.println("----------------------------------------");
                System.out.print("|  " + parser.getRow());
            } else if (SpreadsheetParser.START_CELL == eventType) {
                // System.out.println("开始解析单元格:" + parser.getRow() + "," + parser.getCol() + ":" + parser.getValue());
                System.out.print("  |  " + parser.getValue());
            } else if (SpreadsheetParser.END_CELL == eventType) {
                System.out.print("");
            } else if (SpreadsheetParser.END_RECORD == eventType) {
                System.out.println("  |  ");
            } else if (SpreadsheetParser.END_WORKSHEET == eventType) {
                System.out.println("----------------------------------------");
                System.out.println("Worksheet解析完毕:" + parser.getWorksheetIndex() + "," + parser.getWorksheetName());
            } else if (SpreadsheetParser.END_WORKBOOK == eventType) {
                System.out.println("Workbook解析完毕");
            }
        } while (SpreadsheetParser.EOF != parser.next());

        parser.close();
    }
}
