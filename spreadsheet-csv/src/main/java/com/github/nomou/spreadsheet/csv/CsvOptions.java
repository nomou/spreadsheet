package com.github.nomou.spreadsheet.csv;

import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.nio.charset.Charset;

/**
 * CSV parser / writer options.
 *
 * @author vacoor
 * @since 1.0
 */
class CsvOptions {
    public static final String OPTION_ENCODING_KEY = "encoding";
    public static final String OPTION_SEPARATOR_CHAR_KEY = "separator_char";
    public static final String OPTION_QUOTE_CHAR_KEY = "quote_char";
    public static final String OPTION_ESCAPE_CHAR_KEY = "escape_char";

    private Charset encoding;
    private char separatorChar;
    private char quoteChar;
    private char escapeChar;

    public CsvOptions(final Charset encoding, final char separatorChar, final char quoteChar, final char escapeChar) {
        this.encoding = encoding;
        this.separatorChar = separatorChar;
        this.quoteChar = quoteChar;
        this.escapeChar = escapeChar;
    }

    public CsvParser createParser(final InputStream in) {
        return new CsvParser(new InputStreamReader(in, encoding), separatorChar, quoteChar, escapeChar);
    }

    public CsvWriter createWriter(final OutputStream out) {
        return new CsvWriter(new OutputStreamWriter(out, encoding), separatorChar, quoteChar, escapeChar);
    }


    public CsvOptions set(final String option, final Object value) {
        if (OPTION_ENCODING_KEY.equals(option)) {
            if (value instanceof String) {
                this.setEncoding(Charset.forName((String) value));
            } else if (value instanceof Charset) {
                this.setEncoding((Charset) value);
            } else {
                throw new IllegalArgumentException("illegal option '" + option + "' value '" + value + "'.");
            }
        } else if (OPTION_SEPARATOR_CHAR_KEY.equals(option)) {
            if (value instanceof Character) {
                this.setSeparatorChar((Character) value);
            } else {
                throw new IllegalArgumentException("illegal option '" + option + "' value '" + value + "'.");
            }
        } else if (OPTION_QUOTE_CHAR_KEY.equals(option)) {
            if (value instanceof Character) {
                this.setQuoteChar((Character) value);
            } else {
                throw new IllegalArgumentException("illegal option '" + option + "' value '" + value + "'.");
            }
        } else if (OPTION_ESCAPE_CHAR_KEY.equals(option)) {
            if (value instanceof Character) {
                this.setEscapeChar((Character) value);
            } else {
                throw new IllegalArgumentException("illegal option '" + option + "' value '" + value + "'.");
            }
        }
        return this;
    }

    public Charset getEncoding() {
        return encoding;
    }

    public void setEncoding(final Charset encoding) {
        this.encoding = encoding;
    }

    public char getSeparatorChar() {
        return separatorChar;
    }

    public void setSeparatorChar(final char separatorChar) {
        this.separatorChar = separatorChar;
    }

    public char getQuoteChar() {
        return quoteChar;
    }

    public void setQuoteChar(final char quoteChar) {
        this.quoteChar = quoteChar;
    }

    public char getEscapeChar() {
        return escapeChar;
    }

    public void setEscapeChar(final char escapeChar) {
        this.escapeChar = escapeChar;
    }
}
