package com.example.docx.sample;

import com.example.docx.WordDocument;
import com.example.docx.page.PageSetup;
import com.example.docx.style.Alignment;
import com.example.docx.style.ParagraphStyle;
import com.example.docx.style.TextStyle;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

public final class AnotherSample {

    private static final String FONT = "Arial";
    private static final String RED = "#DB0011";
    private static final String BLACK = "#000000";
    private static final String GREY = "#595959";

    /** The running header strip that sits at the top of every source page. */
    private static final String RUNNING_HEADER = "Free to View ● ESG - Global    9 September 2021";

    private enum Kind { H1, SUB, SUB6, BODY9, BODY6, NUM }

    private record Block(Kind kind, String text) { }

    /** Source page 2: Disclosure appendix. */
    private static final List<Block> PAGE_2 = List.of(

            new Block(Kind.H1, "Lorem ipsum dolor s"),
            new Block(Kind.SUB, "Lorem ipsum dolor sit"),
            new Block(Kind.BODY9, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium tincidunt lacus, ut interdum tellus elementum sed. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices. Posuere cubilia curae mauris viverra diam vitae quam suscipit varius. Aenean commodo ligula eget dolor aenean massa cum sociis natoque. Penatibus et magnis dis parturient montes nascetur ridiculus mus. Donec quam felis ultricies nec pellentesque eu pretium quis sem. Nulla consequat massa quis en"),
            new Block(Kind.SUB, "Lorem ipsum dolor sit"),
            new Block(Kind.BODY9, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamc"),
            new Block(Kind.BODY9, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium tincidunt lacus, ut interdum tellus elementum sed. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices. Posuere cubilia curae mauris viverra diam vitae quam suscipit varius. Aenean commodo ligula eget do"),
            new Block(Kind.BODY9, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium "),
            new Block(Kind.BODY9, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium tincidunt lacus, ut interdum tellus elementum se"),
            new Block(Kind.BODY9, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. N"),
            new Block(Kind.BODY9, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullam"),
            new Block(Kind.BODY9, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dol"),
            new Block(Kind.BODY9, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui offic"));

    /** Source page 3: Additional disclosures. */
    private static final List<Block> PAGE_3 = List.of(

            new Block(Kind.SUB, "Lorem ipsum dolor sit "),
            new Block(Kind.NUM, "Lorem ipsum dolor sit amet, consectetur adipisci"),
            new Block(Kind.NUM, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exe"),
            new Block(Kind.NUM, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium tincidunt lacus, ut interdum tellus elementum sed. Vestibulum ante ipsum primis in faucibus orci lu"),
            new Block(Kind.NUM, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur "));

    /** Source page 4: Disclaimer. */
    private static final List<Block> PAGE_4 = List.of(

            new Block(Kind.H1, "Lorem ipsu"),
            new Block(Kind.SUB6, "Lorem ipsum dolor sit amet, consecte"),
            new Block(Kind.BODY6, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium tincidunt lacus, ut interdum tellus elementum sed. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices. Posuere cubilia curae mauris viverra diam vitae quam suscipit varius. Aenean commodo ligula eget dolor aenean massa cum sociis natoque. Penatibus et magnis dis parturient montes nascetur ridiculus mus. Donec quam felis ultricies nec pellentesque eu pretium quis sem. Nulla consequat massa quis enim donec pede justo fringilla vel. Aliquet nec vulputate eget arcu in enim justo rhoncus ut imperdiet. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium tincidunt lacus, ut interdum tellus elementum sed. Vestibulum ante ipsum "),
            new Block(Kind.SUB6, "Lorem ipsum dolo"),
            new Block(Kind.BODY6, "Lorem ipsum d"),
            new Block(Kind.BODY6, "Lorem ipsum dolor sit a"),
            new Block(Kind.BODY6, "Lorem ipsum dolor sit a"),
            new Block(Kind.BODY6, "Lorem ipsum dolor sit amet,"),
            new Block(Kind.BODY6, "Lorem ipsum dolor sit"),
            new Block(Kind.BODY6, "Lorem ipsum dolor sit amet, co"),
            new Block(Kind.BODY6, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium tincidunt lacus, ut interdum tellus elementum sed. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices. Posuere cubilia curae mauris viverra diam vitae quam suscipit varius. Aenean commodo ligula eget dolor aenean massa cum sociis natoque. Penatibus et magnis dis parturient montes nascetur ridiculus mus. Donec quam felis ultricies nec pellentesque eu pretium quis sem. Nulla consequat massa quis enim donec pede justo fringilla vel. Aliquet nec vulputate eget arcu in enim justo rhoncus ut imperdiet. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium tincidunt lacus, ut interdum tellus elementum sed. Vestibulum ante ipsum primis in faucibus orci luctus et ultrice"),
            new Block(Kind.BODY6, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ulla"),
            new Block(Kind.BODY6, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium tincidunt lacus, ut interdum tellus elementum sed. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices. Posuere cubilia curae mauris viverra diam vitae quam suscipit varius. Aenean commodo ligula eget dolor aenean massa cum sociis natoque. Penatibus et magnis dis parturient montes nascetur ridiculus mus. Donec quam felis ultricies nec pellentesque eu pretium quis sem. Nulla consequat massa quis enim donec pede justo fringilla vel. Aliquet nec vulputate eget arcu in enim justo rhoncus ut imperdiet. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium tincidunt lacus, ut interdum tellus elementum sed. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices. "),
            new Block(Kind.BODY6, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium tincidunt lacus, ut interdum tellus elementum sed. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices. Posuere cubilia curae mauris viverra diam vitae quam suscipit varius. Aenean commodo ligula eget dolor aenean massa cum sociis natoque. Penatibus et magnis dis parturient montes nascetur ridiculus mus. Donec quam felis ultricies nec pellentesque eu pretium quis sem. Nulla consequat massa quis enim donec pede justo fringilla vel. Aliquet nec vulputate eget arcu in enim justo rhoncus ut imperdiet. Lor"),
            new Block(Kind.BODY6, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium tincidunt lacus, ut interdum tellus elementum sed. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices. Posuere cubilia curae mauris viverra diam vitae quam suscipit varius. Aenean commodo ligula eget dolor aenean massa cum sociis natoque. Penatibus et magnis dis parturient montes nascetur ridiculus mus. Donec quam felis ultricies nec pellentesque eu pretium quis sem. Nulla consequat massa quis enim donec pede justo fringilla vel. Aliquet nec vulputate eget arcu in enim justo rhoncus ut imperdiet. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium tincidunt lacus, ut interdum tellus elementum sed. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices. Posuere cubilia curae mauris viverra diam vitae quam suscipit varius. Aenean commodo ligula eget dolor aenean massa cum sociis natoque. Penatibus et magnis dis parturient montes nascetur ridiculus mus. Donec quam felis ultricies nec pellentesque eu pretium quis sem. Nulla consequat massa quis enim donec pede justo fringilla vel. Aliquet nec vulputate eget arcu in enim justo rhoncus ut imperdiet. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea co"),
            new Block(Kind.BODY6, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium tincidunt lacus, ut interdum tellus elementum sed. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices. Posuere cubilia curae mauris viverra diam vitae quam suscipit varius. Aenean commodo ligula eget dolor aenean massa cum sociis natoque. Penatibus et magnis dis parturient montes nascetur ridiculus mus. Donec quam felis ultricies nec pellentesque eu pretium quis sem. Nulla consequat massa quis enim donec pede justo fringilla vel. Aliquet nec vulputate eget arcu in enim justo rhoncus ut imperdiet. Lorem ipsum dolor sit amet, consec"),
            new Block(Kind.BODY6, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium tincidunt lacus, ut interdum tellus elementum sed. Vest"),
            new Block(Kind.BODY6, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatur excepteur sint occaecat cupidatat non proident. Sunt in culpa qui officia deserunt mollit anim id est laborum. Curabitur pretium tincidunt lacus, ut interdum tellus elementum sed. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices. Posuere cubilia curae mauris viverra diam vitae quam suscipit varius. Aenean commodo ligula eget dolor aenean massa cum sociis natoque. Penatibus et magnis dis parturient montes nascetur ridiculus mus. Donec quam felis ultricies nec pellentesque eu pretium quis sem. Nulla consequat massa quis enim donec pede justo fringilla vel. Aliquet nec vulputate eget arcu in enim justo rhoncu"),
            new Block(Kind.BODY6, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. Nisi ut aliquip ex ea commodo consequat duis aute irure dolor. In reprehenderit in voluptate velit esse cillum dolore eu fugiat. Nulla pariatu"),
            new Block(Kind.BODY6, "Lorem ips"));

    private AnotherSample() {
    }

    static void main(String[] ignored) throws IOException {
        var builder = WordDocument.builder()
                .pageSetup(PageSetup.builder().a4().marginsTwips(851).build());

        renderPage(builder, PAGE_2, false);
        renderPage(builder, PAGE_3, true);
        renderPage(builder, PAGE_4, true);

        WordDocument document = builder.build();

        Path target = Path.of("target", "another-sample-%d.docx".formatted(System.currentTimeMillis()));
        Files.createDirectories(target.getParent());
        try (OutputStream out = Files.newOutputStream(target)) {
            document.writeTo(out);
        }
        System.out.println("Wrote " + target.toAbsolutePath());
    }

    /** Emits one source page: its running header (optionally breaking first) then its blocks. */
    private static void renderPage(WordDocument.Builder builder, List<Block> page, boolean pageBreakBefore) {
        builder.paragraph(RUNNING_HEADER, text(8, false, GREY), headerParagraph(pageBreakBefore));
        for (Block block : page) {
            switch (block.kind()) {
                case H1 -> builder.heading(block.text(), text(27, true, RED), headingParagraph());
                case SUB -> builder.heading(block.text(), text(9, true, BLACK), subHeadingParagraph());
                case SUB6 -> builder.heading(block.text(), text(6, true, BLACK), subHeadingParagraph());
                case BODY9, NUM -> builder.paragraph(block.text(), text(9, false, BLACK), bodyParagraph(120));
                case BODY6 -> builder.paragraph(block.text(), text(6, false, BLACK), bodyParagraph(60));
            }
        }
    }

    private static TextStyle text(double sizePt, boolean bold, String color) {
        return TextStyle.builder().font(FONT).sizePt(sizePt).bold(bold).italic(false).color(color).build();
    }

    /** Small grey header line, right-aligned, that starts a new page when asked. */
    private static ParagraphStyle headerParagraph(boolean pageBreakBefore) {
        return ParagraphStyle.builder()
                .spaceBeforeTwips(0)
                .spaceAfterTwips(180)
                .alignment(Alignment.RIGHT)
                .pageBreakBefore(pageBreakBefore)
                .build();
    }

    /** 27 pt section title: kept with the text that follows it. */
    private static ParagraphStyle headingParagraph() {
        return ParagraphStyle.builder()
                .spaceBeforeTwips(60)
                .spaceAfterTwips(120)
                .keepWithNext(true)
                .keepLines(true)
                .build();
    }

    /** Bold sub-heading (Analyst Certification, Issuer of report, ...): glued to its body. */
    private static ParagraphStyle subHeadingParagraph() {
        return ParagraphStyle.builder()
                .spaceBeforeTwips(120)
                .spaceAfterTwips(40)
                .keepWithNext(true)
                .keepLines(true)
                .build();
    }

    /** Justified body paragraph; splits across pages but never strands a single line. */
    private static ParagraphStyle bodyParagraph(int spaceAfterTwips) {
        return ParagraphStyle.builder()
                .spaceBeforeTwips(0)
                .spaceAfterTwips(spaceAfterTwips)
                .alignment(Alignment.JUSTIFY)
                .widowControl(true)
                .build();
    }
}
