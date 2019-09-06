package com.github.nomou.spreadsheet;

/**
 * Configurable spreadsheet object.
 *
 * @author vacoor
 * @since 1.0
 */
public interface SpreadsheetConfigurable<T> {

    /**
     * configure option.
     *
     * @param option the option name
     * @param value the option value
     * @return the configurable object
     */
    T configure(final String option, final Object value);

}
