package com.excelninja.domain.model;

import java.util.Iterator;
import java.util.List;

/**
 * Closeable iterator for streaming Excel chunk reads.
 *
 * <p>Use with try-with-resources when iteration may stop early.
 */
public interface ChunkReader<T> extends Iterator<List<T>>, AutoCloseable {

    @Override
    void close();
}
