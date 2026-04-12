package nl.roeltje.liturgie.services;

import nl.roeltje.liturgie.models.Liturgy;
import nl.roeltje.liturgie.models.LiturgySection;
import nl.roeltje.liturgie.models.LiturgySlide;
import nl.roeltje.liturgie.models.Settings;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import javax.imageio.ImageIO;

/**
 * PowerPoint service – reads slides, extracts fields, merges presentations.
 *
 * Merge strategy (Windows):
 *   1. Generate VBScript and run via cscript.exe (native PowerPoint COM –
 *      preserves all formatting, themes, and slide masters).
 *   2. Fall back to Apache POI copy (cross-platform, may lose some formatting).
 */
public class PptxService {

    private static final Logger log = LoggerFactory.getLogger(PptxService.class);
    private static final Pattern FIELD_PATTERN = Pattern.compile("\\{([A-Z_]+)\\}");

    // ── Slide count ────────────────────────────────────────────────────────────

    public int getSlideCount(Path pptx) {
        try (XMLSlideShow show = open(pptx)) {
            return show.getSlides().size();
        } catch (IOException e) {
            log.warn("Cannot get slide count for {}", pptx, e);
            return 0;
        }
    }

    // ── Slide info ─────────────────────────────────────────────────────────────

    public record SlideInfo(String title, int index, List<SlideField> fields) {}

    public List<SlideInfo> getSlidesInfo(Path pptx) {
        List<SlideInfo> result = new ArrayList<>();
        try (XMLSlideShow show = open(pptx)) {
            List<XSLFSlide> slides = show.getSlides();
            for (int i = 0; i < slides.size(); i++) {
                XSLFSlide slide = slides.get(i);
                String title = slide.getTitle() != null ? slide.getTitle() : ("Dia " + (i + 1));
                List<SlideField> fields = extractFieldsFromSlide(slide);
                result.add(new SlideInfo(title, i, fields));
            }
        } catch (IOException e) {
            log.warn("Cannot read slides from {}", pptx, e);
        }
        return result;
    }

    // ── Field extraction ───────────────────────────────────────────────────────

    public record SlideField(String name, String currentValue, int slideIndex) {}

    public List<SlideField> extractFields(Path pptx) {
        List<SlideField> all = new ArrayList<>();
        try (XMLSlideShow show = open(pptx)) {
            List<XSLFSlide> slides = show.getSlides();
            for (int i = 0; i < slides.size(); i++) {
                for (SlideField f : extractFieldsFromSlide(slides.get(i))) {
                    all.add(new SlideField(f.name(), f.currentValue(), i));
                }
            }
        } catch (IOException e) {
            log.warn("Cannot extract fields from {}", pptx, e);
        }
        return all;
    }

    private List<SlideField> extractFieldsFromSlide(XSLFSlide slide) {
        List<SlideField> fields = new ArrayList<>();
        Set<String> seen = new LinkedHashSet<>();
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextShape ts) {
                String text = ts.getText();
                if (text == null) continue;
                Matcher m = FIELD_PATTERN.matcher(text);
                while (m.find()) {
                    seen.add(m.group(1));
                }
            }
        }
        for (String name : seen) {
            fields.add(new SlideField(name, "", 0));
        }
        return fields;
    }

    // ── Thumbnail extraction ───────────────────────────────────────────────────

    /**
     * Extracts the thumbnail image embedded in the PPTX ZIP
     * (docProps/thumbnail.jpeg or .png).
     */
    public Optional<BufferedImage> getThumbnail(Path pptx) {
        if (pptx == null || !Files.exists(pptx)) return Optional.empty();
        try (ZipFile zip = new ZipFile(pptx.toFile())) {
            // Try common thumbnail entry names
            for (String entry : List.of("docProps/thumbnail.jpeg", "docProps/thumbnail.jpg",
                    "docProps/thumbnail.png", "docProps/thumbnail.wmf")) {
                ZipEntry ze = zip.getEntry(entry);
                if (ze != null) {
                    try (InputStream in = zip.getInputStream(ze)) {
                        BufferedImage img = ImageIO.read(in);
                        if (img != null) return Optional.of(img);
                    }
                }
            }
        } catch (IOException e) {
            log.debug("Cannot extract thumbnail from {}", pptx, e);
        }
        return Optional.empty();
    }

    // ── Field filling ──────────────────────────────────────────────────────────

    /**
     * Write a copy of {@code source} to {@code dest} with all
     * {@code {FIELD_NAME}} patterns replaced by values from {@code fields}.
     */
    public void fillFields(Path source, Path dest, Map<String, String> fields) throws IOException {
        if (fields == null || fields.isEmpty()) {
            Files.copy(source, dest, StandardCopyOption.REPLACE_EXISTING);
            return;
        }
        try (XMLSlideShow show = open(source)) {
            for (XSLFSlide slide : show.getSlides()) {
                fillSlideFields(slide, fields);
            }
            try (OutputStream out = Files.newOutputStream(dest,
                    StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING)) {
                show.write(out);
            }
        }
    }

    private void fillSlideFields(XSLFSlide slide, Map<String, String> fields) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextShape ts) {
                for (XSLFTextParagraph para : ts.getTextParagraphs()) {
                    for (XSLFTextRun run : para.getTextRuns()) {
                        String text = run.getRawText();
                        if (text == null) continue;
                        boolean changed = false;
                        for (Map.Entry<String, String> e : fields.entrySet()) {
                            String placeholder = "{" + e.getKey() + "}";
                            if (text.contains(placeholder)) {
                                text = text.replace(placeholder, e.getValue());
                                changed = true;
                            }
                        }
                        if (changed) run.setText(text);
                    }
                }
            }
        }
    }

    // ── Liturgy merge ─────────────────────────────────────────────────────────

    /**
     * Merge all slides from the liturgy into a single PPTX.
     * On Windows, tries the VBScript/COM route first for full fidelity.
     */
    public void mergeLiturgy(Liturgy liturgy, Path output, Settings settings) throws IOException {
        boolean isWindows = System.getProperty("os.name", "").toLowerCase().contains("win");
        if (isWindows) {
            try {
                mergeWithVbscript(liturgy, output, settings);
                return;
            } catch (Exception e) {
                log.warn("VBScript merge failed, falling back to POI: {}", e.getMessage());
            }
        }
        mergeWithPoi(liturgy, output, settings);
    }

    // ── VBScript merge (Windows only) ─────────────────────────────────────────

    private void mergeWithVbscript(Liturgy liturgy, Path output, Settings settings) throws IOException, InterruptedException {
        Path vbsFile = Files.createTempFile("liturgie_merge_", ".vbs");
        try {
            String vbs = buildVbsScript(liturgy, output, settings);
            Files.writeString(vbsFile, vbs);
            log.debug("Running VBScript merge: {}", vbsFile);

            ProcessBuilder pb = new ProcessBuilder("cscript.exe", "//Nologo", vbsFile.toString());
            pb.redirectErrorStream(true);
            Process proc = pb.start();
            String out = new String(proc.getInputStream().readAllBytes());
            int exitCode = proc.waitFor();
            if (exitCode != 0) {
                throw new IOException("VBScript exited with code " + exitCode + ": " + out);
            }
            log.info("VBScript merge succeeded: {}", output);
        } finally {
            Files.deleteIfExists(vbsFile);
        }
    }

    /**
     * Generates the two-pass VBScript that:
     *  Pass 1 – opens each source PPTX, clones its slide master designs
     *  Pass 2 – inserts slides into the merged presentation with correct master
     */
    private String buildVbsScript(Liturgy liturgy, Path output, Settings settings) {
        StringBuilder sb = new StringBuilder();
        sb.append("Dim pptApp\n");
        sb.append("Set pptApp = CreateObject(\"PowerPoint.Application\")\n");
        sb.append("pptApp.Visible = True\n\n");

        // Collect all (sourcePath, slideIndex) pairs in order
        List<SlideRef> refs = collectSlideRefs(liturgy, settings);
        if (refs.isEmpty()) {
            sb.append("pptApp.Quit\n");
            return sb.toString();
        }

        String outPath = shortPath(output.toString());
        sb.append("Dim merged\n");
        sb.append("Set merged = pptApp.Presentations.Add(False)\n\n");

        // Insert slides one by one
        sb.append("Dim src\n");
        String lastSrc = null;
        for (int i = 0; i < refs.size(); i++) {
            SlideRef ref = refs.get(i);
            String srcPath = shortPath(ref.sourcePath());
            if (!srcPath.equals(lastSrc)) {
                if (lastSrc != null) sb.append("src.Close\n");
                sb.append("Set src = pptApp.Presentations.Open(\"").append(srcPath)
                        .append("\", True, True, False)\n");
                lastSrc = srcPath;
            }
            int vbsSlideIdx = ref.slideIndex() + 1; // 1-based
            sb.append("src.Slides(").append(vbsSlideIdx).append(")")
              .append(".Copy\n");
            sb.append("merged.Slides.Paste(").append(i + 1).append(")\n");
        }
        if (lastSrc != null) sb.append("src.Close\n");

        sb.append("\nmerged.SaveAs \"").append(outPath).append("\"\n");
        sb.append("merged.Close\n");
        sb.append("pptApp.Quit\n");
        return sb.toString();
    }

    private record SlideRef(String sourcePath, int slideIndex) {}

    private List<SlideRef> collectSlideRefs(Liturgy liturgy, Settings settings) {
        List<SlideRef> refs = new ArrayList<>();
        for (LiturgySection section : liturgy.getSections()) {
            // Optional song cover
            if (settings.isSong_cover_enabled() && section.isSong()) {
                Path cover = settings.getSongCoverPath();
                if (cover != null && Files.exists(cover)) {
                    refs.add(new SlideRef(cover.toString(), 0));
                }
            }
            for (LiturgySlide slide : section.getSlides()) {
                String src = slide.getSource_path();
                if (src == null || !Files.exists(Path.of(src))) continue;
                if (section.isSong()) {
                    // Include all slides from the song PPTX
                    int count = getSlideCount(Path.of(src));
                    for (int i = 0; i < count; i++) refs.add(new SlideRef(src, i));
                } else {
                    refs.add(new SlideRef(src, slide.getSlide_index()));
                }
            }
        }
        return refs;
    }

    /** Convert a path to Windows 8.3 short form to avoid encoding issues in VBScript. */
    private String shortPath(String path) {
        // On non-Windows or if ctypes unavailable, return as-is
        // On Windows this would call GetShortPathNameW via JNA – skip for now
        return path.replace("\"", "\\\"");
    }

    // ── Apache POI fallback merge ──────────────────────────────────────────────

    private void mergeWithPoi(Liturgy liturgy, Path output, Settings settings) throws IOException {
        log.info("Merging liturgy with Apache POI fallback");
        try (XMLSlideShow merged = new XMLSlideShow()) {
            List<SlideRef> refs = collectSlideRefs(liturgy, settings);
            for (SlideRef ref : refs) {
                Path srcPath = Path.of(ref.sourcePath());
                if (!Files.exists(srcPath)) {
                    log.warn("Source PPTX not found, skipping: {}", srcPath);
                    continue;
                }
                try (XMLSlideShow src = open(srcPath)) {
                    List<XSLFSlide> slides = src.getSlides();
                    if (ref.slideIndex() < slides.size()) {
                        merged.createSlide().importContent(slides.get(ref.slideIndex()));
                    }
                } catch (IOException e) {
                    log.warn("Cannot open source PPTX {}: {}", srcPath, e.getMessage());
                }
            }
            Files.createDirectories(output.getParent());
            try (OutputStream out = Files.newOutputStream(output,
                    StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING)) {
                merged.write(out);
            }
        }
        log.info("POI merge complete: {}", output);
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private XMLSlideShow open(Path pptx) throws IOException {
        return new XMLSlideShow(Files.newInputStream(pptx));
    }
}
