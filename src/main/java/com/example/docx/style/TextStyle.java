package com.example.docx.style;

import com.example.docx.DocumentGenerationException;
import com.example.docx.Units;
import java.math.BigInteger;
import java.util.Locale;
import java.util.regex.Pattern;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.Color;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;

/**
 * Immutable run formatting (font family, size in points, bold, italic, colour),
 * emitted as direct formatting. Used by both headings and body text.
 *
 * <p>Note: headings do not appear in Word's Navigation pane because no built-in
 * {@code Heading1} style id is referenced. This is a phase-one trade-off for
 * rendering identically regardless of the document stylesheet.
 */
public final class TextStyle {

    private static final Pattern HEX = Pattern.compile("[0-9A-Fa-f]{6}");

    public static final String DEFAULT_FONT = "Calibri Light";
    public static final double DEFAULT_SIZE_PT = 20.0;
    public static final String DEFAULT_COLOR = "1F4E79";

    /** Word's practical ceiling for {@code ST_HpsMeasure}. */
    public static final double MAX_SIZE_PT = 1638.0;

    private final String fontFamily;
    private final double sizePt;
    private final boolean bold;
    private final boolean italic;
    private final String colorHex;

    private TextStyle(Builder b) {
        this.fontFamily = b.fontFamily;
        this.sizePt = b.sizePt;
        this.bold = b.bold;
        this.italic = b.italic;
        this.colorHex = b.colorHex;
    }

    public static Builder builder() {
        return new Builder();
    }

    public static TextStyle defaults() {
        return builder().build();
    }

    /** Body text: Calibri 11 pt, regular, black. */
    public static TextStyle body() {
        return builder().font("Calibri").sizePt(11).bold(false).color("000000").build();
    }

    public String fontFamily() {
        return fontFamily;
    }

    public double sizePt() {
        return sizePt;
    }

    public boolean bold() {
        return bold;
    }

    public boolean italic() {
        return italic;
    }

    /** Bare uppercase {@code RRGGBB}, never prefixed with {@code #}. */
    public String colorHex() {
        return colorHex;
    }

    /** Builds the {@code w:rPr} for this style. */
    public RPr toRPr() {
        ObjectFactory factory = Context.getWmlObjectFactory();
        RPr rPr = factory.createRPr();

        RFonts fonts = factory.createRFonts();
        fonts.setAscii(fontFamily);
        fonts.setHAnsi(fontFamily);
        rPr.setRFonts(fonts);

        int halfPoints = Units.pointsToHalfPoints(sizePt);
        HpsMeasure size = factory.createHpsMeasure();
        size.setVal(BigInteger.valueOf(halfPoints));
        rPr.setSz(size);

        HpsMeasure complexScriptSize = factory.createHpsMeasure();
        complexScriptSize.setVal(BigInteger.valueOf(halfPoints));
        rPr.setSzCs(complexScriptSize);

        if (bold) {
            BooleanDefaultTrue on = factory.createBooleanDefaultTrue();
            on.setVal(Boolean.TRUE);
            rPr.setB(on);
        }
        if (italic) {
            BooleanDefaultTrue on = factory.createBooleanDefaultTrue();
            on.setVal(Boolean.TRUE);
            rPr.setI(on);
        }

        Color color = factory.createColor();
        color.setVal(colorHex);
        rPr.setColor(color);

        return rPr;
    }

    /** Fluent builder. Every field is defaulted; {@code TextStyle.defaults()} is valid alone. */
    public static final class Builder {

        private String fontFamily = DEFAULT_FONT;
        private double sizePt = DEFAULT_SIZE_PT;
        private boolean bold = true;
        private boolean italic = false;
        private String colorHex = DEFAULT_COLOR;

        private Builder() {
        }

        public Builder font(String fontFamily) {
            this.fontFamily = fontFamily;
            return this;
        }

        public Builder sizePt(double sizePt) {
            this.sizePt = sizePt;
            return this;
        }

        public Builder bold(boolean bold) {
            this.bold = bold;
            return this;
        }

        public Builder italic(boolean italic) {
            this.italic = italic;
            return this;
        }

        /** Accepts {@code #RRGGBB} or {@code RRGGBB}, any case. */
        public Builder color(String hex) {
            this.colorHex = hex;
            return this;
        }

        public TextStyle build() {
            if (fontFamily == null || fontFamily.isBlank()) {
                throw new DocumentGenerationException("font family must not be blank");
            }
            if (!(sizePt > 0) || !Double.isFinite(sizePt)) {
                throw new DocumentGenerationException(
                        "font size must be greater than zero points, got " + sizePt);
            }
            if (sizePt > MAX_SIZE_PT) {
                throw new DocumentGenerationException(
                        "font size must not exceed " + MAX_SIZE_PT + " points, got " + sizePt);
            }
            this.fontFamily = fontFamily.trim();
            this.colorHex = normaliseColour(colorHex);
            return new TextStyle(this);
        }

        private static String normaliseColour(String hex) {
            if (hex == null) {
                throw new DocumentGenerationException("colour must not be null");
            }
            String bare = hex.startsWith("#") ? hex.substring(1) : hex;
            if (!HEX.matcher(bare).matches()) {
                throw new DocumentGenerationException(
                        "colour must be 6 hex digits, optionally prefixed with '#', got '" + hex + "'");
            }
            return bare.toUpperCase(Locale.ROOT);
        }
    }
}
