package com.example.docx.style;

import com.example.docx.DocumentGenerationException;

/** Immutable table appearance: header and body borders, and the text style of each. */
public final class TableStyle {

    private final TableBorderStyle headerBorder;
    private final TableBorderStyle bodyBorder;
    private final TextStyle headerText;
    private final TextStyle bodyText;

    private TableStyle(Builder b) {
        this.headerBorder = b.headerBorder;
        this.bodyBorder = b.bodyBorder;
        this.headerText = b.headerText;
        this.bodyText = b.bodyText;
    }

    public static Builder builder() {
        return new Builder();
    }

    /** Bold header underlined in dark blue, hairline rules between body rows. */
    public static TableStyle defaults() {
        return builder().build();
    }

    public TableBorderStyle headerBorder() {
        return headerBorder;
    }

    public TableBorderStyle bodyBorder() {
        return bodyBorder;
    }

    public TextStyle headerText() {
        return headerText;
    }

    public TextStyle bodyText() {
        return bodyText;
    }

    /** Fluent builder. Every field is defaulted; {@code TableStyle.defaults()} is valid alone. */
    public static final class Builder {

        private TableBorderStyle headerBorder = TableBorderStyle.bottomOnly("#1F4E79", 1.0);
        private TableBorderStyle bodyBorder = TableBorderStyle.bottomOnly("#BFBFBF", 0.5);
        private TextStyle headerText =
                TextStyle.builder().font("Calibri").sizePt(11).bold(true).color("000000").build();
        private TextStyle bodyText = TextStyle.body();

        private Builder() {
        }

        public Builder headerBorder(TableBorderStyle headerBorder) {
            this.headerBorder = requireNonNull(headerBorder, "header border");
            return this;
        }

        public Builder bodyBorder(TableBorderStyle bodyBorder) {
            this.bodyBorder = requireNonNull(bodyBorder, "body border");
            return this;
        }

        public Builder headerText(TextStyle headerText) {
            this.headerText = requireNonNull(headerText, "header text style");
            return this;
        }

        public Builder bodyText(TextStyle bodyText) {
            this.bodyText = requireNonNull(bodyText, "body text style");
            return this;
        }

        public TableStyle build() {
            return new TableStyle(this);
        }

        private static <T> T requireNonNull(T value, String field) {
            if (value == null) {
                throw new DocumentGenerationException(field + " must not be null");
            }
            return value;
        }
    }
}
