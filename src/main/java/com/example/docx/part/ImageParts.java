package com.example.docx.part;

import com.example.docx.DocumentGenerationException;
import com.example.docx.Units;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.StringReader;
import javax.imageio.ImageIO;
import javax.xml.parsers.DocumentBuilderFactory;
import org.docx4j.dml.CTBlip;
import org.docx4j.dml.CTOfficeArtExtension;
import org.docx4j.dml.CTOfficeArtExtensionList;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.contenttype.ContentType;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart.AddPartBehaviour;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.w3c.dom.Element;
import org.xml.sax.InputSource;

/**
 * Builds image paragraphs.
 *
 * <p>Unlike {@code content} and {@code style}, this needs the package: an image is a
 * part plus a relationship, so it cannot be a pure function of its arguments.
 *
 * <p>An SVG is stored the way Word stores one — the blip's {@code r:embed} points at a
 * PNG, and the SVG rides along as an {@code asvg:svgBlip} extension. Renderers that
 * understand the extension draw the vector; the rest draw the PNG they already had.
 */
public final class ImageParts {

    /** Identifies the SVG blip extension. Defined by Microsoft; do not change. */
    private static final String SVG_EXTENSION_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

    private static final String ASVG_NS =
            "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
    private static final String RELATIONSHIPS_NS =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private ImageParts() {
    }

    /**
     * A paragraph holding the SVG, drawn {@code widthTwips} wide with the height taken
     * from the PNG's aspect ratio.
     */
    public static P svgImage(WordprocessingMLPackage pkg, byte[] svg, byte[] png, int widthTwips) {
        if (pkg == null) {
            throw new DocumentGenerationException("package must not be null");
        }
        if (svg == null || svg.length == 0) {
            throw new DocumentGenerationException("svg bytes must not be empty");
        }
        if (png == null || png.length == 0) {
            throw new DocumentGenerationException("png fallback bytes must not be empty");
        }
        if (widthTwips <= 0) {
            throw new DocumentGenerationException(
                    "image width must be greater than zero twips, got " + widthTwips);
        }

        BufferedImage raster = decodePng(png);
        MainDocumentPart mainDocumentPart = pkg.getMainDocumentPart();

        try {
            // The PNG is the blip's primary image: the fallback every renderer understands.
            BinaryPartAbstractImage pngPart = BinaryPartAbstractImage.createImagePart(pkg, png);
            Inline inline = pngPart.createImageInline("image", "Image", 1, 2, false);

            // docx4j has no ImageSvgPart, and its part factory treats image/svg+xml as
            // XML and hands back a DefaultXmlPart, so the SVG part is built by hand.
            BinaryPart svgPart = new BinaryPart(new PartName("/word/media/image.svg"));
            svgPart.setBinaryData(svg);
            svgPart.setContentType(new ContentType("image/svg+xml"));
            svgPart.setRelationshipType(Namespaces.IMAGE);
            Relationship svgRelationship = mainDocumentPart.addTargetPart(
                    svgPart, AddPartBehaviour.RENAME_IF_NAME_EXISTS);

            attachSvgExtension(inline, svgRelationship.getId());
            setExtent(inline, widthTwips, raster.getWidth(), raster.getHeight());

            ObjectFactory factory = Context.getWmlObjectFactory();
            Drawing drawing = factory.createDrawing();
            drawing.getAnchorOrInline().add(inline);
            R run = factory.createR();
            run.getContent().add(drawing);
            P paragraph = factory.createP();
            paragraph.getContent().add(run);
            return paragraph;
        } catch (DocumentGenerationException e) {
            throw e;
        } catch (Exception e) {
            throw new DocumentGenerationException("failed to embed the image", e);
        }
    }

    private static BufferedImage decodePng(byte[] png) {
        try {
            BufferedImage image = ImageIO.read(new ByteArrayInputStream(png));
            if (image == null) {
                throw new DocumentGenerationException("png fallback is not a readable image");
            }
            return image;
        } catch (IOException e) {
            throw new DocumentGenerationException("failed to read the png fallback", e);
        }
    }

    private static void attachSvgExtension(Inline inline, String svgRelationshipId)
            throws Exception {
        CTBlip blip = inline.getGraphic().getGraphicData().getPic().getBlipFill().getBlip();

        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        // Mandatory, and false by default. Left off, the prefixed name parses as a
        // literal element in NO namespace and Word silently ignores the SVG.
        factory.setNamespaceAware(true);
        // The string parsed here is fully library-constructed today (fixed namespaces
        // plus a docx4j-generated relationship id), so this isn't reachable yet — but
        // hardening now means a future edit that threads less-trusted input through this
        // factory doesn't reintroduce XXE by accident.
        factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
        factory.setFeature("http://xml.org/sax/features/external-general-entities", false);
        factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
        factory.setXIncludeAware(false);
        factory.setExpandEntityReferences(false);
        Element svgBlip = factory.newDocumentBuilder()
                .parse(new InputSource(new StringReader(
                        "<asvg:svgBlip xmlns:asvg=\"" + ASVG_NS
                                + "\" xmlns:r=\"" + RELATIONSHIPS_NS
                                + "\" r:embed=\"" + svgRelationshipId + "\"/>")))
                .getDocumentElement();

        CTOfficeArtExtension extension = new CTOfficeArtExtension();
        extension.setUri(SVG_EXTENSION_URI);
        extension.setAny(svgBlip);

        CTOfficeArtExtensionList extensions = new CTOfficeArtExtensionList();
        extensions.getExt().add(extension);
        blip.setExtLst(extensions);
    }

    private static void setExtent(Inline inline, int widthTwips, int pngWidthPx, int pngHeightPx) {
        // createImageInline sizes from the PNG's intrinsic dimensions and DPI, which is
        // not what the caller asked for, so the extent is set explicitly.
        long cx = Units.twipsToEmu(widthTwips);
        long cy = Math.round(cx * (double) pngHeightPx / pngWidthPx);
        inline.getExtent().setCx(cx);
        inline.getExtent().setCy(cy);
    }
}
