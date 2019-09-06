package com.github.nomou.spreadsheet.csv;

/**
 * Copyright 2005 Bytecode Pty Ltd.
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * http://www.apache.org/licenses/LICENSE-2.0
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import java.io.Closeable;
import java.io.IOException;
import java.io.LineNumberReader;
import java.io.Reader;

/**
 * A very simple CSV reader released under a commercial-friendly license.
 *
 * @author Glen Smith
 */
class CsvParser implements Closeable {

    private final LineNumberReader lnr;

    private boolean hasNext = true;

    CsvLineParser parser;

    int skipLines;

    private boolean linesSkiped;

    /**
     * The default line to start reading.
     */
    public static final int DEFAULT_SKIP_LINES = 0;

    /**
     * Constructs CsvParser using a comma for the separator.
     *
     * @param reader the reader to an underlying CSV source.
     */
    public CsvParser(Reader reader) {
        this(reader, CsvLineParser.DEFAULT_SEPARATOR, CsvLineParser.DEFAULT_QUOTE_CHARACTER, CsvLineParser.DEFAULT_ESCAPE_CHARACTER);
    }

    /**
     * Constructs CsvParser with supplied separator.
     *
     * @param reader    the reader to an underlying CSV source.
     * @param separator the delimiter to use for separating entries.
     */
    public CsvParser(Reader reader, char separator) {
        this(reader, separator, CsvLineParser.DEFAULT_QUOTE_CHARACTER, CsvLineParser.DEFAULT_ESCAPE_CHARACTER);
    }

    /**
     * Constructs CsvParser with supplied separator and quote char.
     *
     * @param reader    the reader to an underlying CSV source.
     * @param separator the delimiter to use for separating entries
     * @param quotechar the character to use for quoted elements
     */
    public CsvParser(Reader reader, char separator, char quotechar) {
        this(reader, separator, quotechar, CsvLineParser.DEFAULT_ESCAPE_CHARACTER, DEFAULT_SKIP_LINES, CsvLineParser.DEFAULT_STRICT_QUOTES);
    }

    /**
     * Constructs CsvParser with supplied separator, quote char and quote handling
     * behavior.
     *
     * @param reader       the reader to an underlying CSV source.
     * @param separator    the delimiter to use for separating entries
     * @param quotechar    the character to use for quoted elements
     * @param strictQuotes sets if characters outside the quotes are ignored
     */
    public CsvParser(Reader reader, char separator, char quotechar, boolean strictQuotes) {
        this(reader, separator, quotechar, CsvLineParser.DEFAULT_ESCAPE_CHARACTER, DEFAULT_SKIP_LINES, strictQuotes);
    }

    /**
     * Constructs CsvParser with supplied separator and quote char.
     *
     * @param reader    the reader to an underlying CSV source.
     * @param separator the delimiter to use for separating entries
     * @param quotechar the character to use for quoted elements
     * @param escape    the character to use for escaping a separator or quote
     */

    public CsvParser(Reader reader, char separator, char quotechar, char escape) {
        this(reader, separator, quotechar, escape, DEFAULT_SKIP_LINES, CsvLineParser.DEFAULT_STRICT_QUOTES);
    }

    /**
     * Constructs CsvParser with supplied separator and quote char.
     *
     * @param reader    the reader to an underlying CSV source.
     * @param separator the delimiter to use for separating entries
     * @param quotechar the character to use for quoted elements
     * @param line      the line number to skip for start reading
     */
    public CsvParser(Reader reader, char separator, char quotechar, int line) {
        this(reader, separator, quotechar, CsvLineParser.DEFAULT_ESCAPE_CHARACTER, line, CsvLineParser.DEFAULT_STRICT_QUOTES);
    }

    /**
     * Constructs CsvParser with supplied separator and quote char.
     *
     * @param reader    the reader to an underlying CSV source.
     * @param separator the delimiter to use for separating entries
     * @param quotechar the character to use for quoted elements
     * @param escape    the character to use for escaping a separator or quote
     * @param line      the line number to skip for start reading
     */
    public CsvParser(Reader reader, char separator, char quotechar, char escape, int line) {
        this(reader, separator, quotechar, escape, line, CsvLineParser.DEFAULT_STRICT_QUOTES);
    }

    /**
     * Constructs CsvParser with supplied separator and quote char.
     *
     * @param reader       the reader to an underlying CSV source.
     * @param separator    the delimiter to use for separating entries
     * @param quotechar    the character to use for quoted elements
     * @param escape       the character to use for escaping a separator or quote
     * @param line         the line number to skip for start reading
     * @param strictQuotes sets if characters outside the quotes are ignored
     */
    public CsvParser(Reader reader, char separator, char quotechar, char escape, int line, boolean strictQuotes) {
        this(reader, separator, quotechar, escape, line, strictQuotes, CsvLineParser.DEFAULT_IGNORE_LEADING_WHITESPACE);
    }

    /**
     * Constructs CsvParser with supplied separator and quote char.
     *
     * @param reader                  the reader to an underlying CSV source.
     * @param separator               the delimiter to use for separating entries
     * @param quotechar               the character to use for quoted elements
     * @param escape                  the character to use for escaping a separator or quote
     * @param line                    the line number to skip for start reading
     * @param strictQuotes            sets if characters outside the quotes are ignored
     * @param ignoreLeadingWhiteSpace it true, parser should ignore white space before a quote in a field
     */
    public CsvParser(Reader reader, char separator, char quotechar, char escape, int line, boolean strictQuotes, boolean ignoreLeadingWhiteSpace) {
        this(reader, line, new CsvLineParser(separator, quotechar, escape, strictQuotes, ignoreLeadingWhiteSpace));
    }

    /**
     * Constructs CsvParser with supplied separator and quote char.
     *
     * @param reader        the reader to an underlying CSV source.
     * @param line          the line number to skip for start reading
     * @param lineParser the parser to use to parse input
     */
    public CsvParser(Reader reader, int line, CsvLineParser lineParser) {
        this.lnr = (reader instanceof LineNumberReader ? (LineNumberReader) reader : new LineNumberReader(reader));
        this.skipLines = line;
        this.parser = lineParser;
    }

    public int getLineNumber() {
        return lnr.getLineNumber();
    }

    public boolean hasNext() {
        return hasNext;
    }

    /**
     * Reads the next line from the buffer and converts to a string array.
     *
     * @return a string array with each comma-separated element as a separate
     * entry.
     * @throws IOException if bad things happen during the read
     */
    public String[] next() throws IOException {
        String[] result = null;
        do {
            String nextLine = readNextLine();
            if (!hasNext) {
                return result; // should throw if still pending?
            }
            String[] r = parser.parseLineMulti(nextLine);
            if (r.length > 0) {
                if (result == null) {
                    result = r;
                } else {
                    String[] t = new String[result.length + r.length];
                    System.arraycopy(result, 0, t, 0, result.length);
                    System.arraycopy(r, 0, t, result.length, r.length);
                    result = t;
                }
            }
        } while (parser.isPending());
        return result;
    }

    /**
     * Reads the next line from the file.
     *
     * @return the next line from the file without trailing newline
     * @throws IOException if bad things happen during the read
     */
    private String readNextLine() throws IOException {
        if (!this.linesSkiped) {
            for (int i = 0; i < skipLines; i++) {
                lnr.readLine();
            }
            this.linesSkiped = true;
        }
        String nextLine = lnr.readLine();
        if (nextLine == null) {
            hasNext = false;
        }
        return hasNext ? nextLine : null;
    }

    /**
     * Closes the underlying reader.
     *
     * @throws IOException if the close fails
     */
    public void close() throws IOException {
        lnr.close();
    }
}