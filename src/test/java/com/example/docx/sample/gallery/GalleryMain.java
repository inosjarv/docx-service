package com.example.docx.sample.gallery;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.local.office.LocalOfficeManager;

/**
 * Renders every {@link GalleryFixtures fixture} through LibreOffice and writes one
 * {@code target/gallery/index.html} so they can all be reviewed in one browser tab,
 * without hand-running each sample main and opening its output one file at a time.
 *
 * <p>Requires a local LibreOffice install (a standard Homebrew-cask install on macOS is
 * auto-detected; see DOCUMENTATION.md). Run with
 * {@code JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn test-compile exec:java@gallery}.
 */
public final class GalleryMain {

    private static final Path GALLERY_DIR = Path.of("target", "gallery");

    private GalleryMain() {
    }

    public static void main(String[] args) throws IOException {
        OfficeManager officeManager;
        try {
            // Catching the broader RuntimeException here (not just OfficeException) is
            // deliberate: JODConverter's builder().build() throws unchecked when
            // LibreOffice can't be found (e.g. NullPointerException when auto-detection
            // finds nothing, IllegalStateException for an invalid configured path),
            // before officeManager.start() ever gets a chance to throw the checked
            // OfficeException. Don't narrow this back to OfficeException alone.
            officeManager = LocalOfficeManager.builder().install().build();
            officeManager.start();
        } catch (OfficeException | RuntimeException e) {
            System.err.println("Could not start LibreOffice: " + e.getMessage());
            System.err.println("Install it (e.g. `brew install --cask libreoffice` on "
                    + "macOS) and try again. See DOCUMENTATION.md's \"Document gallery\" "
                    + "section.");
            System.exit(1);
            return;
        }

        try {
            List<RenderedFixture> rendered = new ArrayList<>();
            for (GalleryFixture fixture : GalleryFixtures.all()) {
                rendered.add(render(officeManager, fixture));
            }

            Files.createDirectories(GALLERY_DIR);
            Path indexPath = GALLERY_DIR.resolve("index.html");
            Files.writeString(indexPath, GalleryIndexWriter.write(rendered));
            System.out.println("Wrote " + indexPath.toAbsolutePath());
        } finally {
            try {
                officeManager.stop();
            } catch (OfficeException e) {
                System.err.println("Failed to stop LibreOffice cleanly: " + e.getMessage());
            }
        }
    }

    /**
     * One fixture's build+render, isolated so a single bad fixture (a build that throws,
     * a conversion LibreOffice chokes on) shows up as a visible failure note in the
     * index rather than aborting every other fixture's output. Catching the broad
     * {@code Exception} here is deliberate: a fixture's {@code builder()} can throw an
     * unchecked {@code DocumentGenerationException}, and {@code DocxToImages} throws the
     * checked {@code OfficeException}/{@code IOException} — both must be caught the same
     * way, without aborting the run.
     */
    private static RenderedFixture render(OfficeManager officeManager, GalleryFixture fixture) {
        try {
            List<Path> pages = DocxToImages.renderPages(
                    officeManager, fixture.builder().get(), GALLERY_DIR.resolve(fixture.name()));
            List<Path> relativePages = pages.stream()
                    .map(page -> Path.of(fixture.name()).resolve(page.getFileName()))
                    .toList();
            return RenderedFixture.success(fixture.name(), relativePages);
        } catch (Exception e) {
            System.err.println("Failed to render '" + fixture.name() + "': " + e);
            e.printStackTrace();
            return RenderedFixture.failure(fixture.name(), e.toString());
        }
    }
}
