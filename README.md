# Spreadsheet - streaming spreadsheet API
  电子表格(csv, microsoft excel) 流式API（适用于不考虑表格样式,主要考虑数据的场景）
  
# Usage
### 写入
```java
    final OutputStream out = new FileOutputStream("out.xlsx");
    final SpreadsheetWriter writer = Spreadsheet.getWriterFactory("xlsx").create(out);
    try {
        writer.configure("encoding", "UTF-8")
                .start()
                .write("SKU编码", "商品名称", "UPC编码", "价格").next()
                .write("B10-4").write("B10-4").write("B10-4").write(35.4);
    } finally {
        writer.close();
    }
```

### 按照事件解析
```java
    final InputStream in = new BufferedInputStream(new FileInputStream("out.xlsx"));
    final SpreadsheetParser parser = Spreadsheet.getParserFactory("xlsx", "xls", "csv").create(in);
    parser.configure("encoding", "UTF-8");

    try {
        int next;
        while (parser.hasNext()) {
            next = parser.next();
            if (SpreadsheetParser.START_WORKBOOK == next) {
                // TODO
                System.out.println("Start parsing workbook");
            } else if (SpreadsheetParser.END_WORKBOOK == next) {
                System.out.println("End of workbook");
            } else if (SpreadsheetParser.START_WORKSHEET == next) {
                System.out.println("Start parsing worksheet");
            } else if (SpreadsheetParser.END_WORKSHEET == next) {
                System.out.println("End of worksheet");
            } else if (SpreadsheetParser.START_RECORD == next) {
                System.out.println("start parsing record, row=" + parser.getRow());
            } else if (SpreadsheetParser.END_RECORD == next) {
                System.out.println("end of record, row=" + parser.getRow());
            } else if (SpreadsheetParser.START_CELL == next) {
                final int row = parser.getRow();
                final int col = parser.getCol();
                final Object value = parser.getValue();
                System.out.println(String.format("start cell, row=%s, col=%s, value=%s", row, col, value));
            } else if (SpreadsheetParser.END_CELL == next) {
                final int row = parser.getRow();
                final int col = parser.getCol();
                final Object value = parser.getValue();
                System.out.println(String.format("end cell, row=%s, col=%s, value=%s", row, col, value));
            } else {
                System.out.println("Other ---------: " + next);
            }
        }

    } finally {
        parser.close();
    }
```

### 按照行解析
```java
    final InputStream in = new BufferedInputStream(new FileInputStream("out.xlsx"));
    final SpreadsheetParser parser = Spreadsheet.getParserFactory("xlsx", "xls", "csv").create(in);
    try {
        parser.configure("encoding", "UTF-8");

        Object[] record;
        do {
            record = parser.nextRecord(true);
            if (null != record) {
                final int row = parser.getRow();
                final SpreadsheetRecord r = SpreadsheetRecord.wrap(record);
                System.out.println(row + ":" + Arrays.toString(record) + "/" + r.getString("A"));
            }
        } while (null != record);
    } finally {
        parser.close();
    }
```