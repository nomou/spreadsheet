package com.github.nomou.spreadsheet;

import com.github.nomou.spreadsheet.spi.SpreadsheetParserFactory;
import com.github.nomou.spreadsheet.spi.SpreadsheetWriterFactory;
import com.github.nomou.spreadsheet.util.SpreadsheetUtils;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.ServiceLoader;

/**
 * Spreadsheet parser/writer locator.
 *
 * @author changhe.yang
 * @since 20190807
 */
public final class Spreadsheet {
    private static final List<Format> ALL_FORMATS = new LinkedList<>();
    private static final Map<byte[], List<SpreadsheetParserFactory>> PARSER_FACTORIES_MAP;
    private static final Map<String, List<SpreadsheetWriterFactory>> WRITER_FACTORIES_MAP;


    static {
        PARSER_FACTORIES_MAP = loadSpreadsheetParserFactories();
        WRITER_FACTORIES_MAP = loadSpreadsheetWriterFactories();
    }

    /**
     * Spreadsheet format.
     */
    public static class Format {
        private final String name;
        private final byte[] header;
        private final String[] extensions;

        public Format(final String name, final byte[] header, final String... extensions) {
            this.name = name;
            this.header = header;
            this.extensions = extensions;
        }

        @Override
        public boolean equals(final Object o) {
            if (this == o) {
                return true;
            }
            if (o == null || getClass() != o.getClass()) {
                return false;
            }

            final Format format = (Format) o;

            if (!Arrays.equals(header, format.header)) {
                return false;
            }
            // Probably incorrect - comparing Object[] arrays with Arrays.equals
            return Arrays.equals(extensions, format.extensions);
        }

        @Override
        public int hashCode() {
            int result = Arrays.hashCode(header);
            result = 31 * result + Arrays.hashCode(extensions);
            return result;
        }

        @Override
        public String toString() {
            return "Format {" + name + "(" + Arrays.toString(extensions) + ") }";
        }
    }

    /**
     * Get a spreadsheet writer factory that supports a given extension.
     *
     * @param extension the extension
     * @return the spreadsheet writer factory
     */
    public static SpreadsheetWriterFactory getWriterFactory(final String extension) {
        final List<SpreadsheetWriterFactory> factories = WRITER_FACTORIES_MAP.get(extension.toLowerCase());
        if (null == factories || factories.isEmpty()) {
            throw new IllegalStateException("No suite SpreadsheetWriterFactory found for " + extension);
        }
        return factories.iterator().next();
    }

    /**
     * Get a spreadsheet parser factory that supports a given extensions.
     *
     * <p>if 'extensions' is not specified, the returned factory will support all supported formats.</p>
     *
     * @param extensions the extensions
     * @return the spreadsheet parser factory
     */
    public static SpreadsheetParserFactory getParserFactory(final String... extensions) {
        List<Format> formats;
        Map<byte[], List<SpreadsheetParserFactory>> factoryMap;
        if (0 == extensions.length) {
            // using all
            formats = getFormatsByHeaders(PARSER_FACTORIES_MAP.keySet());
            factoryMap = PARSER_FACTORIES_MAP;
        } else {
            final Map<byte[], List<SpreadsheetParserFactory>> filtered = new HashMap<>();
            formats = getFormatsByExtensions(extensions);
            for (final Format format : formats) {
                filtered.put(format.header, PARSER_FACTORIES_MAP.get(format.header));
            }
            factoryMap = filtered;
        }

        if (factoryMap.isEmpty()) {
            throw new IllegalStateException("No suite SpreadsheetParserFactory found for " + Arrays.toString(extensions));
        }
        final Format[] formatArray = new Format[formats.size()];
        return new MixedSpreadsheetParserFactory(formats.toArray(formatArray), factoryMap);
    }

    /* ************************************
     *
     * ********************************** */

    private static List<Format> getFormatsByExtensions(final String... extensions) {
        final List<Format> formats = new ArrayList<>(extensions.length);
        for (final Format format : ALL_FORMATS) {
            for (final String ext : extensions) {
                if (Arrays.asList(format.extensions).contains(ext)) {
                    formats.add(format);
                }
            }
        }
        return formats;
    }

    private static List<Format> getFormatsByHeaders(final Collection<byte[]> headers) {
        final List<Format> formats = new ArrayList<>(headers.size());
        for (final byte[] header : headers) {
            final Format format = getFormatByHeader(header);
            if (null != format) {
                formats.add(format);
            }
        }
        return formats;
    }

    private static Format getFormatByHeader(final byte[] header) {
        for (final Format format : ALL_FORMATS) {
            if (Arrays.equals(format.header, header)) {
                return format;
            }
        }
        return null;
    }

    /* ************************************
     *
     * ********************************** */

    /**
     * Loads all spreadsheet parser factories.
     *
     * @return the spreadsheet parser factories, map.key: file header bytes, map.value: file-header-bytes-factories
     */
    private static Map<byte[], List<SpreadsheetParserFactory>> loadSpreadsheetParserFactories() {
        final Map<byte[], List<SpreadsheetParserFactory>> factoriesMap = new HashMap<>();
        final ServiceLoader<SpreadsheetParserFactory> loader = ServiceLoader.load(SpreadsheetParserFactory.class);
        for (final SpreadsheetParserFactory factory : loader) {
            final Format[] formats = factory.getSupportedFormats();
            for (final Format format : formats) {
                if (!ALL_FORMATS.contains(format)) {
                    ALL_FORMATS.add(format);
                }

                final byte[] header = format.header;
                List<SpreadsheetParserFactory> factories = factoriesMap.get(header);
                if (null == factories) {
                    factories = new ArrayList<>();
                    factoriesMap.put(header, factories);
                }
                factories.add(factory);
            }
        }
        return factoriesMap;
    }

    /**
     * Loads all spreadsheet writer factories.
     *
     * @return the spreadsheet writer factories, map.key: extensions, map.value: extension-factories
     */
    private static Map<String, List<SpreadsheetWriterFactory>> loadSpreadsheetWriterFactories() {
        final Map<String, List<SpreadsheetWriterFactory>> factoriesMap = new HashMap<>(10);
        final ServiceLoader<SpreadsheetWriterFactory> loader = ServiceLoader.load(SpreadsheetWriterFactory.class);
        for (final SpreadsheetWriterFactory factory : loader) {
            final Format[] formats = factory.getSupportedFormats();
            for (final Format format : formats) {
                if (!ALL_FORMATS.contains(format)) {
                    ALL_FORMATS.add(format);
                }
                for (final String extension : format.extensions) {
                    List<SpreadsheetWriterFactory> factories = factoriesMap.get(extension.toLowerCase());
                    if (null == factories) {
                        factories = new ArrayList<>();
                        factoriesMap.put(extension.toLowerCase(), factories);
                    }
                    factories.add(factory);
                }
            }
        }
        return factoriesMap;
    }


    /* *****************************************
     *
     * *************************************** */

    /**
     * Mixed spreadsheet parser factory.
     */
    private static class MixedSpreadsheetParserFactory implements SpreadsheetParserFactory {
        private final Format[] formats;
        private final Map<byte[], List<SpreadsheetParserFactory>> factoriesMap;

        private MixedSpreadsheetParserFactory(final Format[] formats,
                                              final Map<byte[], List<SpreadsheetParserFactory>> factories) {
            this.formats = formats;
            this.factoriesMap = factories;
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public Format[] getSupportedFormats() {
            return formats;
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public SpreadsheetParser create(final InputStream in) throws SpreadsheetException {
            try {
                List<SpreadsheetParserFactory> factories = SpreadsheetUtils.matches(in, factoriesMap);
                if (null != factories) {
                    if (factories.isEmpty()) {
                        throw new SpreadsheetException("file type error");
                    }

                    SpreadsheetException lastError = null;
                    for (final SpreadsheetParserFactory factory : factories) {
                        try {
                            return factory.create(in);
                        } catch (final SpreadsheetException ex) {
                            lastError = ex;
                        }
                    }

                    throw lastError;
                } else {
                    // unsupported 'mark' method, try first factory.
                    for (final List<SpreadsheetParserFactory> tryFactories : factoriesMap.values()) {
                        for (final SpreadsheetParserFactory tryFactory : tryFactories) {
                            try {
                                return tryFactory.create(in);
                            } catch (final SpreadsheetException ex) {
                                // 'in.markSupported()' always 'false'
                                if (!in.markSupported()) {
                                    throw new SpreadsheetException("input must be available markSupported,you can do like this 'new BufferedInputStream(new FileInputStream(\"/xxxx\"))'");
                                }
                            }
                        }
                    }

                    throw new SpreadsheetException("not supported");
                }
            } catch (IOException e) {
                throw new SpreadsheetException(e);
            }
        }
    }

    /**
     * Private constructor.
     */
    private Spreadsheet() {
    }

}
