package com.github.nomou.spreadsheet.util;

import freework.util.Castor;

import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 * Spreadsheet record(row).
 *
 * @author vacoor
 * @since 1.0
 */
public class SpreadsheetRecord {
    /**
     * The record cells data.
     */
    private final List<?> data;

    /**
     * Create a record instance for the given data.
     *
     * @param data the cells data
     * @return the record
     */
    public static SpreadsheetRecord wrap(final Object[] data) {
        return wrap(null != data ? Arrays.asList(data) : null);
    }

    /**
     * Create a record instance for the given data.
     *
     * @param data the cells data
     * @return the record
     */
    public static SpreadsheetRecord wrap(final List<?> data) {
        return new SpreadsheetRecord(data);
    }

    /**
     * Create a record instance for the given data.
     *
     * @param data the cells data
     */
    private SpreadsheetRecord(final List<?> data) {
        this.data = data;
    }

    public boolean isNull() {
        return null == data;
    }

    public boolean isEmpty() {
        return null == data || data.isEmpty();
    }

    /**
     * Get the number of cells.
     *
     * @return the number of cells
     */
    public int size() {
        return !isNull() ? data.size() : 0;
    }

    /**
     * Get the cell value for the given index of column.
     *
     * @param columnIndex the index of column
     * @return the cell value
     */
    public Object getValue(final int columnIndex) {
        return columnIndex < size() ? data.get(columnIndex) : null;
    }

    /**
     * Get the cell value for the given column name.
     *
     * @param columnName the name of column
     * @return the cell value
     */
    public Object getValue(final String columnName) {
        return getValue(nameToIndex(columnName));
    }

    /**
     * Get the cell boolean value for the given index of column.
     *
     * @param columnIndex the index of column
     * @return the cell value
     */
    public Boolean getBoolean(final int columnIndex) {
        return columnIndex < size() ? Castor.asBoolean(data.get(columnIndex)) : null;
    }

    /**
     * Get the cell boolean value for the given column name.
     *
     * @param columnName the name of column
     * @return the cell value
     */
    public Boolean getBoolean(final String columnName) {
        return getBoolean(nameToIndex(columnName));
    }

    /**
     * Get the cell number value for the given index of column.
     *
     * @param columnIndex the index of column
     * @return the cell value
     */
    public Number getNumber(final int columnIndex) {
        return columnIndex < size() ? Castor.asNumber(data.get(columnIndex), Number.class) : null;
    }

    /**
     * Get the cell number value for the given column name.
     *
     * @param columnName the name of column
     * @return the cell value
     */
    public Number getNumber(final String columnName) {
        return getNumber(nameToIndex(columnName));
    }

    /**
     * Get the cell string value for the given index of column.
     *
     * @param columnIndex the index of column
     * @return the cell value
     */
    public String getString(final int columnIndex) {
        return columnIndex < size() ? Castor.asString(data.get(columnIndex)) : null;
    }

    /**
     * Get the cell string value for the given column name.
     *
     * @param columnName the name of column
     * @return the cell value
     */
    public String getString(final String columnName) {
        return getString(nameToIndex(columnName));
    }

    /**
     * Get the cell date value for the given index of column.
     *
     * @param columnIndex the index of column
     * @return the cell value
     */
    public Date getDate(final int columnIndex) {
        return columnIndex < size() ? Castor.asDate(data.get(columnIndex)) : null;
    }

    /**
     * Get the cell date value for the given column name.
     *
     * @param columnName the name of column
     * @return the cell value
     */
    public Date getDate(final String columnName) {
        return getDate(nameToIndex(columnName));
    }

    private int nameToIndex(final String columnName) {
        int column = -1;
        for (int i = 0; i < columnName.length(); ++i) {
            final int c = columnName.charAt(i);
            column = (column + 1) * 26 + c - 'A';
        }
        return column;
    }
}
