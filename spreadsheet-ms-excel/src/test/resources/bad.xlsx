# Streaming data API for Spreadsheet.
```
final InputSource in = ...;
final SpreadsheetParser parser = new OpenXMLSpreadsheetParser(in);

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
} while (parser.hasNext() && SpreadsheetParser.EOF != (eventType = parser.next()));
```