package com.github.nomou.spreadsheet.msexcel;

import com.github.nomou.spreadsheet.AbstractSpreadsheetParser;
import com.github.nomou.spreadsheet.SpreadsheetException;
import com.github.nomou.spreadsheet.SpreadsheetParser;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipTypes;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument;
import org.xml.sax.SAXException;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import javax.xml.stream.events.XMLEvent;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;

/**
 * POI-based spreadsheet parser for Microsoft Excel 2007+.
 *
 * @author vacoor
 * @see <a href="http://poi.apache.org/">POI</a>
 * @since 1.0
 */
class OpenXMLSpreadsheetParser extends AbstractSpreadsheetParser {
    /* *********************
     *     POI Objects.
     * ******************* */
    /**
     * SpreadsheetML Package.
     */
    private OPCPackage spreadsheet;

    /**
     * Workbook Part.
     */
    private PackagePart workbookPart;

    /**
     * Workbook shared string table.
     */
    private ReadOnlySharedStringsTable sharedStringsTable;

    /**
     * Workbook style table.
     */
    private StylesTable stylesTable;

    /**
     * Worksheets summary.
     */
    private CTSheet[] worksheets;

    /**
     * Current worksheet part.
     */
    private PackagePart worksheetPart;

    /**
     * Current worksheet parser.
     */
    private OpenXMLWorksheetParser worksheetParser = null;

    public OpenXMLSpreadsheetParser(final InputStream in) throws SpreadsheetException {
        initInputSource(in);
    }

    @Override
    public SpreadsheetParser configure(final String option, final Object value) {
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getNumberOfWorksheets() {
        return null != this.worksheets ? this.worksheets.length : 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getRow() {
        final int st = eventType;
        if (START_RECORD != st && END_RECORD != st && START_CELL != st && END_CELL != st) {
            throw new IllegalStateException("getRow() called in illegal state");
        }
        return this.worksheetParser.row;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getCol() {
        final int st = this.eventType;
        if (START_CELL != st && END_CELL != st) {
            throw new IllegalStateException("getCol() called in illegal state");
        }
        return this.worksheetParser.col;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Object getValue() {
        final int st = this.eventType;
        if (START_CELL != st && END_CELL != st) {
            throw new IllegalStateException("getValue() called in illegal state");
        }
        return this.worksheetParser.value;
    }

    /* **********************************************************
     *
     * ******************************************************** */

    /**
     * {@inheritDoc}
     */
    @Override
    protected int doNext() throws SpreadsheetException {
        final int t = this.eventType;
        final int worksheets = this.worksheets.length;
        final int worksheetIndex = this.worksheetIndex;

        try {
            int newEvent = EOF;
            if (START_WORKBOOK == t) {
                if (0 < worksheets) {
                    this.worksheetIndex = 0;
                    doPreParseWorksheet();
                    newEvent = START_WORKSHEET;
                } else {
                    doPostParseWorkbook();
                    newEvent = END_WORKBOOK;
                }
            } else if (END_WORKSHEET == t) {
                if (worksheetIndex < worksheets - 1) {
                    this.worksheetIndex++;
                    doPreParseWorksheet();
                    newEvent = START_WORKSHEET;
                } else {
                    doPostParseWorkbook();
                    newEvent = END_WORKBOOK;
                }
            } else if (START_WORKSHEET == t || START_RECORD == t || END_RECORD == t || START_CELL == t || END_CELL == t) {
                final int innerEvent = worksheetParser.next();
                if (END_WORKSHEET == innerEvent) {
                    doPostParseWorksheet();
                }
                newEvent = innerEvent;
            }

            return newEvent;
        } catch (final InvalidFormatException e) {
            throw new SpreadsheetException(e.getMessage(), e.getCause());
        } catch (final IOException e) {
            throw new SpreadsheetException(e);
        } catch (final XMLStreamException e) {
            throw new SpreadsheetException(e);
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void close() throws SpreadsheetException {
        if (END_WORKBOOK != this.eventType) {
            this.eventType = EOF;
        }
        try {
            doPostParseWorkbook();
        } catch (final IOException e) {
            throw new SpreadsheetException(e);
        }
    }

    /**
     * Configure input source for the parser.
     * <p>supports input source type: InputStream, File, String(absolute path).
     *
     * @param inputSource input source.
     * @throws IllegalArgumentException if this input source not support
     * @throws SpreadsheetException     If this input source is parsed incorrectly
     */
    public void initInputSource(final Object inputSource) throws SpreadsheetException {
        try {
            OPCPackage spreadsheet;
            if (inputSource instanceof InputStream) {
                spreadsheet = OPCPackage.open((InputStream) inputSource);
            } else if (inputSource instanceof File) {
                spreadsheet = OPCPackage.open((File) inputSource, PackageAccess.READ);
            } else if (inputSource instanceof String) {
                spreadsheet = OPCPackage.open((String) inputSource, PackageAccess.READ);
            } else {
                throw new IllegalArgumentException("Unsupported input source: " + inputSource);
            }

            this.spreadsheet = spreadsheet;
            this.doPreParseWorkbook();
            this.eventType = START_WORKBOOK;
        } catch (final InvalidFormatException e) {
            throw new SpreadsheetException(e.getMessage(), e.getCause());
        } catch (final IOException e) {
            throw new SpreadsheetException(e);
        } catch (final XmlException e) {
            throw new SpreadsheetException(e);
        } catch (final SAXException e) {
            throw new SpreadsheetException(e);
        }
    }

    /* ******************************************************
     *                   APACHE POI OPERATIONS
     * **************************************************** */

    /**
     * TODO doc me.
     */
    protected Object toJavaObject(final String t, final String s, final String value) {
        Object parsed = null;
        if ("b".equals(t)) {
            // boolean
            parsed = 1 > value.length() || value.charAt(0) == '0' ? Boolean.FALSE : Boolean.TRUE;
        } else if ("e".equals(t)) {
            // error
            parsed = "ERROR:" + value;
        } else if ("inlineStr".equals(t)) {
            // TODO: have seen an example of this, so it's untested.
            final XSSFRichTextString rtsi = new XSSFRichTextString(value);
            parsed = rtsi.toString();
        } else if ("s".equals(t)) {
            // sstindex
            try {
                final int idx = Integer.parseInt(value);
                final XSSFRichTextString rtss = new XSSFRichTextString(sharedStringsTable.getEntryAt(idx));
                parsed = rtss.toString();
            } catch (final NumberFormatException ex) {
                throw new IllegalStateException("Failed to toJavaObject SST index '" + value + "': " + ex.toString());
            }
        } else if ("str".equals(t)) {
            // A formula could result in a string value,
            // so always add double-quote characters.
            parsed = value;
        } else if ("n".equals(t)) {
            parsed = Double.parseDouble(value);
        } else if (null != s) {
            // It's a number, but almost certainly one
            // with a special style or format
            final int styleIndex = Integer.parseInt(s);
            final XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);

            final short formatIndex = style.getDataFormat();
            String formatString = style.getDataFormatString();

            if (null == formatString) {
                formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
            }

            final String n = value;
            // 判断是否是日期格式
            if (HSSFDateUtil.isADateFormat(formatIndex, n)) {
                final Double d = Double.parseDouble(n);
                parsed = HSSFDateUtil.getJavaDate(d);
            } else if (formatString != null) {
                parsed = new DataFormatter().formatRawCellContents(Double.parseDouble(n), formatIndex, formatString);
            } else {
                parsed = n;
            }
        } else if (null == t) {
            // cellType = null as string
            // null == t and null == s, FIXME: 1.0245899999999
            parsed = value;
        } else {
            throw new IllegalStateException("(TODO: Unexpected type: " + t + ")");
        }

        return parsed;
    }

    /**
     * TODO doc me.
     */
    void doPreParseWorkbook() throws IOException, SAXException, XmlException {
        final OPCPackage spreadsheet = this.spreadsheet;

        // new XSSFReader(spreadsheet);
        final List<PackagePart> docParts = spreadsheet.getPartsByRelationshipType(PackageRelationshipTypes.CORE_DOCUMENT);
        final PackagePart workbookPart = 0 < docParts.size() ? docParts.get(0) : null;
        if (null == workbookPart) {
            throw new IllegalStateException("Can not be found CORE_DOCUMENT from the input source");
        }

        final ReadOnlySharedStringsTable sharedStringsTable = new ReadOnlySharedStringsTable(spreadsheet);
        final List<PackagePart> styleParts = spreadsheet.getPartsByContentType(XSSFRelation.STYLES.getContentType());
        final PackagePart stylePart = 0 < styleParts.size() ? styleParts.get(0) : null;
        final StylesTable stylesTable = null != stylePart ? new StylesTable(stylePart, null) : null;
        if (null != stylePart) {
            final List<PackagePart> themeParts = spreadsheet.getPartsByContentType(XSSFRelation.THEME.getContentType());
            if (0 < themeParts.size()) {
                stylesTable.setTheme(new ThemesTable(themeParts.get(0), null));
            }
        }

        final CTWorkbook workbook = WorkbookDocument.Factory.parse(workbookPart.getInputStream()).getWorkbook();
        final List<CTSheet> sheets = workbook.getSheets().getSheetList();

        this.workbookPart = workbookPart;
        this.sharedStringsTable = sharedStringsTable;
        this.stylesTable = stylesTable;
        this.worksheets = sheets.toArray(new CTSheet[sheets.size()]);
    }

    /**
     * TODO doc me.
     */
    void doPostParseWorkbook() throws IOException {
        // reset worksheet.
        this.worksheetIndex = -1;
        this.worksheetName = null;

        closeQuiet(this.workbookPart);

        this.worksheets = new CTSheet[0];
        this.stylesTable = null;
        this.sharedStringsTable = null;
        this.workbookPart = null;

        if (null != this.spreadsheet) {
            this.spreadsheet.close();
        }
        this.spreadsheet = null;
    }

    /**
     * TODO doc me.
     */
    void doPreParseWorksheet() throws XMLStreamException, IOException, InvalidFormatException {
        final CTSheet[] worksheets = this.worksheets;
        final PackagePart workbookPart = this.workbookPart;
        final int worksheetIndex = this.worksheetIndex;
        final String id = worksheets[worksheetIndex].getId();

        final PackageRelationship rel = workbookPart.getRelationship(id);
        final PackagePartName relName = PackagingURIHelper.createPartName(rel.getTargetURI());
        final PackagePart worksheetPart = workbookPart.getPackage().getPart(relName);
        final InputStream in = worksheetPart.getInputStream();

        this.worksheetName = worksheets[worksheetIndex].getName();
        this.worksheetPart = worksheetPart;
        this.worksheetParser = new OpenXMLWorksheetParser(in);
    }

    /**
     * TODO doc me.
     */
    void doPostParseWorksheet() throws IOException, XMLStreamException {
        closeQuiet(this.worksheetParser);
        closeQuiet(this.worksheetPart);

        this.worksheetParser = null;
        this.worksheetPart = null;

        // reset worksheet, move to doPostParseWorkbook.
        // this.worksheetIndex = -1;
        // this.worksheetName = null;
    }

    /**
     * TODO doc me.
     */
    void closeQuiet(final PackagePart part) {
        try {
            if (null != part) {
                part.close();
            }
        } catch (final Exception e) {
            // ignore.
        }
    }

    /**
     * TODO doc me.
     */
    void closeQuiet(final OpenXMLWorksheetParser worksheetParser) {
        try {
            if (null != worksheetParser) {
                worksheetParser.close();
            }
        } catch (final XMLStreamException e) {
            // ignore.
        }
    }

    /* *************************************************************
     *           SpreadsheetML WORKSHEET PARSER.
     * *********************************************************** */

    /**
     * SpreadsheetML worksheet parser.
     */
    private class OpenXMLWorksheetParser {
        private final String ROW_TAG = "row";
        private final String CELL_TAG = "c";
        private final List<String> VALUE_TAGS = Arrays.asList("v", "t");
        private final String TYPE_ATTRIBUTE = "t";
        private final String STYLE_ATTRIBUTE = "s";

        /**
         * worksheet StAX reader.
         */
        private XMLStreamReader reader;

        /**
         * Current parsing state.
         */
        private int state;

        /**
         * Current row.
         */
        private int row = -1;

        /**
         * Current cell column.
         */
        private int col = -1;

        /**
         * Current cell value.
         */
        private Object value;

        /**
         * Create a SpreadsheetML worksheet parser using given inputstream.
         *
         * @param in worksheet inputstream.
         * @throws XMLStreamException
         */
        OpenXMLWorksheetParser(final InputStream in) throws XMLStreamException {
            this.reader = XMLInputFactory.newFactory().createXMLStreamReader(in);
            this.state = START_WORKSHEET;
        }

        /**
         * TODO doc me.
         */
        void close() throws XMLStreamException {
            if (null != this.reader) {
                this.reader.close();
            }
        }

        /**
         * TODO doc me.
         */
        int next() throws XMLStreamException {
            final int st = this.state;
            final XMLStreamReader reader = this.reader;

            int newEvent = EOF;
            if (START_WORKSHEET == st) {
                newEvent = nextRowOpened(reader, true) ? START_RECORD : END_WORKSHEET;
            } else if (START_RECORD == st || END_CELL == st) {
                newEvent = nextCellOpened(reader) ? START_CELL : (nextRowClosed(reader) ? END_RECORD : EOF);
            } else if (START_CELL == st) {
                newEvent = nextCellClosed(reader) ? END_CELL : EOF;
            } else if (END_RECORD == st) {
                newEvent = nextRowOpened(reader, false) ? START_RECORD : END_WORKSHEET;
            }

            return this.state = newEvent;
        }

        /* ************************************
         *
         * ********************************** */

        /**
         * TODO doc me.
         */
        boolean nextRowOpened(final XMLStreamReader reader, boolean allowSkip) throws XMLStreamException {
            if (nextOpenedTag(reader, ROW_TAG, allowSkip)) {
                final String r = reader.getAttributeValue(null, "r");
                try {
                    this.row = Integer.valueOf(r) - 1;
                } catch (final NumberFormatException nfe) {
                    throw new IllegalStateException("element 'row' attribute 'r' is not a valid value");
                }
                return true;
            }

            this.row = -1;
            return false;
        }

        /**
         * TODO doc me.
         */
        boolean nextRowClosed(final XMLStreamReader reader) throws XMLStreamException {
            // this.row = -1;
            return nextClosedTag(reader, ROW_TAG, false);
        }

        /**
         * TODO doc me.
         */
        boolean nextCellOpened(final XMLStreamReader reader) throws XMLStreamException {
            if (nextOpenedTag(reader, CELL_TAG, false)) {
                final String r = reader.getAttributeValue(null, "r");
                this.col = parseColumn(r);

                nextCellValue(reader);

                return true;
            }
            this.col = -1;
            this.value = null;
            return false;
        }

        /**
         * TODO doc me.
         */
        boolean nextCellClosed(final XMLStreamReader reader) throws XMLStreamException {
            /*-
             * <c><v>text</v></c>
             * <c><is><t>text</t></is></c>
             */
            // this.col = -1;
            return nextClosedTag(reader, CELL_TAG, true);
        }

        /**
         * Parse current cell element value.
         *
         * @param reader StAX reader.
         * @return found?
         * @throws XMLStreamException parse exception.
         */
        boolean nextCellValue(final XMLStreamReader reader) throws XMLStreamException {
            boolean found = false;
            String text = null;
            int event = reader.getEventType();
            String localName = XMLEvent.START_ELEMENT == event ? reader.getLocalName() : null;
            if (!CELL_TAG.equals(localName)) {
                throw new IllegalStateException("must be in start element 'c'");
            }

            final String t = reader.getAttributeValue(null, TYPE_ATTRIBUTE);
            final String s = reader.getAttributeValue(null, STYLE_ATTRIBUTE);
            while (reader.hasNext()) {
                event = reader.next();

                // parse 'v'/'t' element content(value).
                if (XMLEvent.START_ELEMENT == event && VALUE_TAGS.contains(localName = reader.getLocalName())) {
                    while (reader.hasNext()) {
                        event = reader.next();
                        if (XMLEvent.END_ELEMENT == event) {
                            break;
                        }
                        if (XMLEvent.CHARACTERS == event || XMLEvent.CDATA == event || XMLEvent.SPACE == event || XMLEvent.ENTITY_REFERENCE == event) {
                            found = true;
                            text = null != text ? text + reader.getText() : reader.getText();
                        }
                    }
                    break;
                }

                // found cell closed tag.
                if (XMLEvent.END_ELEMENT == event && CELL_TAG.equals(localName)) {
                    break;
                }
            }
            this.value = asJavaObject(t, s, text);
            return found;
        }

        /**
         * TODO doc me.
         */
        int parseColumn(final String r) {
            int firstDigit = -1;
            for (int c = 0; c < r.length(); ++c) {
                if (Character.isDigit(r.charAt(c))) {
                    firstDigit = c;
                    break;
                }
            }
            return nameToColumn(r.substring(0, firstDigit));
        }

        /**
         * Converts an Excel column name like "C" to a zero-based index.
         *
         * @param name column index name.
         * @return Index corresponding to the specified name
         */
        int nameToColumn(String name) {
            int column = -1;
            for (int i = 0; i < name.length(); ++i) {
                int c = name.charAt(i);
                column = (column + 1) * 26 + c - 'A';
            }
            return column;
        }

        /**
         * TODO doc me.
         */
        Object asJavaObject(final String t, final String s, final String value) {
            return toJavaObject(t, s, value);
        }

        /* ************************************
         *
         * ********************************** */

        /**
         * Parse until given element open tag.
         *
         * @param reader    StAX reader.
         * @param localName tagName.
         * @return found?
         * @throws XMLStreamException parse exception.
         */
        boolean nextOpenedTag(final XMLStreamReader reader, final String localName, final boolean allowSkip) throws XMLStreamException {
            int event = reader.getEventType();
            String tagName = XMLEvent.START_ELEMENT == event ? reader.getLocalName() : null;

            while (!localName.equals(tagName) && reader.hasNext()) {
                event = reader.next();
                tagName = XMLEvent.START_ELEMENT == event ? reader.getLocalName() : null;
                if (!allowSkip && (XMLEvent.END_ELEMENT == event || XMLEvent.START_ELEMENT == event)) {
                    break;
                }
            }
            return localName.equals(tagName);
        }

        /**
         * Parse until given element closed tag.
         *
         * @param reader    StAX reader.
         * @param localName tagName.
         * @return found?
         * @throws XMLStreamException parse exception.
         */
        boolean nextClosedTag(final XMLStreamReader reader, final String localName, final boolean allowSkip) throws XMLStreamException {
            int event = reader.getEventType();
            String tagName = XMLEvent.END_ELEMENT == event ? reader.getLocalName() : null;

            while (!localName.equals(tagName) && reader.hasNext()) {
                event = reader.next();
                tagName = XMLEvent.END_ELEMENT == event ? reader.getLocalName() : null;
                if (!allowSkip && (XMLEvent.END_ELEMENT == event || XMLEvent.START_ELEMENT == event)) {
                    break;
                }
            }
            return localName.equals(tagName);
        }
    }
}
