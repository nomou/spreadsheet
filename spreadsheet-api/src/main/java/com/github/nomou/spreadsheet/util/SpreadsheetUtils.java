package com.github.nomou.spreadsheet.util;

import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * Spreadsheet utils.
 *
 * @author vacoor
 * @since 1.0
 */
public class SpreadsheetUtils {
    /**
     * Length-first comparator.
     */
    private static final Comparator<byte[]> LEN_FIRST_COMPARATOR = new Comparator<byte[]>() {
        @Override
        public int compare(final byte[] o1, final byte[] o2) {
            return Long.compare(o2.length, o1.length);
        }
    };

    /**
     * Private constructor.
     */
    private SpreadsheetUtils() {
    }

    /**
     * Gets all matched factories based-on file header.
     *
     * @param in           the input source
     * @param candidateMap the candidate factories
     * @param <T>          the factory type
     * @return null if the input source not supported 'mark' method, otherwise all matches results
     * @throws IOException if an I/O error occurs
     */
    public static <T> List<T> matches(final InputStream in, final Map<byte[], List<T>> candidateMap) throws IOException {
        if (!in.markSupported()) {
            return null;
        }

        final byte[] bytes = peekFirstNBytes(in, 8);
        final List<T> candidates = candidateMap.get(bytes);
        if (null != candidates) {
            return candidates;
        }

        final List<T> ret = new LinkedList<>();
        final List<T> any = new LinkedList<>();

        // length first.
        final List<byte[]> keys = new ArrayList<>(candidateMap.keySet());
        Collections.sort(keys, LEN_FIRST_COMPARATOR);
        for (final byte[] expect : keys) {
            if (null == expect || 0 >= expect.length) {
                any.addAll(candidateMap.get(expect));
                continue;
            }

            final byte[] given = bytes.length > expect.length ? Arrays.copyOf(bytes, expect.length) : bytes;
            if (Arrays.equals(given, expect)) {
                ret.addAll(candidateMap.get(expect));
            }
        }
        ret.addAll(any);
        return ret;
    }

    /**
     * Peeks the first N bytes of a given markable stream.
     *
     * @param in    the markable stream
     * @param limit N bytes
     * @return the first limit bytes
     * @throws IOException if an I/O error occurs
     */
    public static byte[] peekFirstNBytes(final InputStream in, final int limit) throws IOException {
        in.mark(limit);

        final byte[] bytes = new byte[limit];
        final int len = in.read(bytes);

        if (len == 0) {
            throw new IllegalStateException("empty stream");
        }

        if (in instanceof PushbackInputStream) {
            ((PushbackInputStream) in).unread(bytes, 0, len);
        } else {
            in.reset();
        }
        return bytes;
    }
}
