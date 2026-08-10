package com.example.docx;

/**
 * Conversions into the units OOXML actually stores.
 *
 * <p>Page geometry is in twips (1/1440 inch); font sizes are in half-points.
 * These helpers convert caller-supplied values only. Built-in constants such as
 * A4's dimensions are literals elsewhere, because deriving them rounds wrong.
 */
public final class Units {

    private static final double TWIPS_PER_INCH = 1440.0;
    private static final double CM_PER_INCH = 2.54;

    private Units() {
    }

    public static int cmToTwips(double cm) {
        check(cm, "centimetres");
        return roundHalfUp(cm / CM_PER_INCH * TWIPS_PER_INCH);
    }

    public static int inchesToTwips(double inches) {
        check(inches, "inches");
        return roundHalfUp(inches * TWIPS_PER_INCH);
    }

    public static int pointsToHalfPoints(double points) {
        check(points, "points");
        return roundHalfUp(points * 2.0);
    }

    private static void check(double value, String unit) {
        if (!Double.isFinite(value)) {
            throw new DocumentGenerationException(unit + " must be a finite number, got " + value);
        }
        if (value < 0) {
            throw new DocumentGenerationException(unit + " must not be negative, got " + value);
        }
    }

    private static int roundHalfUp(double value) {
        return Math.toIntExact(Math.round(value));
    }
}
