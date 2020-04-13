package com.github.nomou.spreadsheet.msexcel;

import com.github.nomou.spreadsheet.AbstractSpreadsheetParser;
import com.github.nomou.spreadsheet.SpreadsheetException;
import com.github.nomou.spreadsheet.SpreadsheetParser;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipTypes;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.xmlbeans.XmlException;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import javax.xml.stream.events.XMLEvent;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * POI-based spreadsheet parser for Microsoft Excel 2007+.
 *
 * @author vacoor
 * @see <a href="http://poi.apache.org/">POI</a>
 * @see XSSFReader
 * @see org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler
 * @since 1.0
 */
class OpenXMLSpreadsheetParser extends AbstractSpreadsheetParser {
    private static final Set<String> WORKSHEET_RELS = Collections.unmodifiableSet(
            new HashSet<String>(Arrays.asList(XSSFRelation.WORKSHEET.getRelation(), XSSFRelation.CHARTSHEET.getRelation()))
    );

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
    private SharedStringsTable sharedStringsTable;

    /**
     * Workbook style table.
     */
    private StylesTable stylesTable;

    /**
     * Worksheets summary.
     */
    private XSSFReader.SheetIterator worksheets;

    /**
     * Worksheets rel-id : worksheet part name.
     */
    private Map<String, PackagePartName> worksheetRelNames;

    /**
     * Current worksheet parser.
     */
    private OpenXMLWorksheetParser worksheetParser = null;

    /**
     * Creates a Open-XML spreadsheet parser.
     *
     * @param in the Open-XML spreadhseet stream
     * @throws SpreadsheetException
     */
    public OpenXMLSpreadsheetParser(final InputStream in) throws SpreadsheetException {
        this.initInputSource(in);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public SpreadsheetParser configure(final String option, final Object value) {
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getNumberOfWorksheets() {
        return worksheetRelNames.size();
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
        try {
            int newEvent = EOF;
            if (START_WORKBOOK == t || END_WORKSHEET == t) {
                if (worksheets.hasNext()) {
                    this.doPreParseWorksheet();
                    newEvent = START_WORKSHEET;
                } else {
                    this.doPostParseWorkbook();
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
            throw new SpreadsheetException("XML Stream error", e);
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
            this.doPostParseWorkbook();
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
            } else if (null != formatString && null != n) {
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
    void doPreParseWorkbook() throws IOException, XmlException {
        try {
            final PackageRelationship coreDocRelationship = spreadsheet.getRelationshipsByType(PackageRelationshipTypes.CORE_DOCUMENT).getRelationship(0);
            if (null == coreDocRelationship) {
                if (null != spreadsheet.getRelationshipsByType(PackageRelationshipTypes.STRICT_CORE_DOCUMENT).getRelationship(0)) {
                    throw new XmlException("Strict OOXML isn't currently supported, please see POI bug #57699");
                }
                throw new XmlException("OOXML file structure broken/invalid - no core document found!");
            }
            this.workbookPart = spreadsheet.getPart(coreDocRelationship);

            /*-
             * StyleTable 不同版本构造器不一样, 因此这里改用 XSSFReader 来读取.
             */
            final XSSFReader reader = new XSSFReader(spreadsheet);
            this.stylesTable = reader.getStylesTable();
            this.sharedStringsTable = reader.getSharedStringsTable();
            this.worksheets = (XSSFReader.SheetIterator) reader.getSheetsData();

            final Map<String, PackagePartName> sheetRelNameMap = new HashMap<String, PackagePartName>();
            for (final PackageRelationship rel : workbookPart.getRelationships()) {
                String relType = rel.getRelationshipType();
                if (WORKSHEET_RELS.contains(relType)) {
                    sheetRelNameMap.put(rel.getId(), PackagingURIHelper.createPartName(rel.getTargetURI()));
                }
            }
            this.worksheetRelNames = sheetRelNameMap;
        } catch (final OpenXML4JException ex) {
            throw new XmlException(ex);
        }
    }

    /**
     * TODO doc me.
     */
    void doPostParseWorkbook() throws IOException {
        // reset worksheet.
        this.worksheetIndex = -1;
        this.worksheetName = null;
        this.worksheets = null;
        this.worksheetRelNames = null;

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
        final InputStream worksheetIn = this.worksheets.next();

        this.worksheetIndex++;
        this.worksheetName = this.worksheets.getSheetName();
        this.worksheetParser = new OpenXMLWorksheetParser(worksheetIn);
    }

    /**
     * TODO doc me.
     */
    void doPostParseWorksheet() throws IOException, XMLStreamException {
        closeQuiet(this.worksheetParser);
//        closeQuiet(this.worksheetPart);

        this.worksheetParser = null;
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
