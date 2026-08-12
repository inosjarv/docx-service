package com.example.docx.part;

import com.example.docx.DocumentGenerationException;
import com.example.docx.content.Hyperlink;
import com.example.docx.content.Paragraphs;
import com.example.docx.style.ParagraphStyle;
import com.example.docx.style.TextStyle;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;

/**
 * Builds a paragraph whose trailing content is a clickable link.
 *
 * <p>Unlike {@code content.Paragraphs}, this needs the package: a hyperlink's target is
 * a relationship ({@code word/_rels/document.xml.rels}, {@code TargetMode="External"}),
 * not inline text, and a relationship cannot exist before the package does.
 */
public final class Hyperlinks {

    private Hyperlinks() {
    }

    /**
     * A paragraph made of an optional leading run and a trailing hyperlink.
     *
     * @param leadingText the body text before the link; blank or null omits the
     *                    leading run entirely, leaving a link-only paragraph
     */
    public static P paragraph(WordprocessingMLPackage pkg, String leadingText,
                              TextStyle leadingStyle, ParagraphStyle paragraphStyle,
                              Hyperlink link) {
        if (pkg == null) {
            throw new DocumentGenerationException("package must not be null");
        }
        if (leadingStyle == null) {
            throw new DocumentGenerationException("text style must not be null");
        }
        if (paragraphStyle == null) {
            throw new DocumentGenerationException("paragraph style must not be null");
        }
        if (link == null) {
            throw new DocumentGenerationException("link must not be null");
        }

        ObjectFactory factory = Context.getWmlObjectFactory();
        P paragraph = factory.createP();
        paragraph.setPPr(paragraphStyle.toPPr());

        if (leadingText != null && !leadingText.isBlank()) {
            paragraph.getContent().add(Paragraphs.run(leadingText, leadingStyle));
        }
        paragraph.getContent().add(hyperlinkRun(pkg, link, factory));
        return paragraph;
    }

    private static org.docx4j.wml.P.Hyperlink hyperlinkRun(
            WordprocessingMLPackage pkg, Hyperlink link, ObjectFactory factory) {
        MainDocumentPart mainDocumentPart = pkg.getMainDocumentPart();
        RelationshipsPart relationshipsPart = mainDocumentPart.getRelationshipsPart(true);

        Relationship relationship = new org.docx4j.relationships.ObjectFactory().createRelationship();
        relationship.setType(Namespaces.HYPERLINK);
        relationship.setTarget(link.url());
        relationship.setTargetMode("External");
        relationshipsPart.addRelationship(relationship);

        org.docx4j.wml.P.Hyperlink hyperlink = new org.docx4j.wml.P.Hyperlink();
        hyperlink.setId(relationship.getId());
        hyperlink.setHistory(true);

        R run = Paragraphs.run(link.text(), link.style());
        hyperlink.getContent().add(run);
        return hyperlink;
    }
}
