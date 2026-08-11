package com.example.docx;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

class UnitsTest {

    @Test
    void cmToTwips() {
        assertEquals(850, Units.cmToTwips(1.5));
        assertEquals(1440, Units.cmToTwips(2.54));
        assertEquals(0, Units.cmToTwips(0.0));
    }

    @Test
    void inchesToTwips() {
        assertEquals(1440, Units.inchesToTwips(1.0));
        assertEquals(720, Units.inchesToTwips(0.5));
    }

    @Test
    void pointsToHalfPoints() {
        assertEquals(40, Units.pointsToHalfPoints(20));
        assertEquals(23, Units.pointsToHalfPoints(11.5));
    }

    @Test
    void rejectsNonFiniteInput() {
        assertThrows(DocumentGenerationException.class, () -> Units.cmToTwips(Double.NaN));
        assertThrows(DocumentGenerationException.class, () -> Units.inchesToTwips(Double.POSITIVE_INFINITY));
    }

    @Test
    void rejectsNegativeInput() {
        assertThrows(DocumentGenerationException.class, () -> Units.cmToTwips(-1.0));
        assertThrows(DocumentGenerationException.class, () -> Units.pointsToHalfPoints(-1.0));
    }

    @Test
    void oversizedConversionsThrowTheModuleException() {
        assertThrows(DocumentGenerationException.class, () -> Units.cmToTwips(1e9));
        assertThrows(DocumentGenerationException.class, () -> Units.inchesToTwips(1e9));
        assertThrows(DocumentGenerationException.class, () -> Units.pointsToHalfPoints(1e12));
    }
}
