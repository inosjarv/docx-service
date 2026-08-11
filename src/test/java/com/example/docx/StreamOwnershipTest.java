package com.example.docx;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.style.HeadingStyle;
import java.io.ByteArrayOutputStream;
import org.junit.jupiter.api.Test;

class StreamOwnershipTest {

    /** ByteArrayOutputStream.close() is a no-op, so track the call explicitly. */
    private static final class TrackingStream extends ByteArrayOutputStream {
        boolean closed;

        @Override
        public void close() {
            closed = true;
        }
    }

    @Test
    void writeToLeavesTheCallersStreamOpen() {
        WordDocument doc = WordDocument.builder()
                .heading("Title", HeadingStyle.defaults())
                .build();

        TrackingStream out = new TrackingStream();
        doc.writeTo(out);

        assertTrue(out.size() > 0, "should have written bytes");
        assertFalse(out.closed, "writeTo must not close the caller's stream");
    }
}
