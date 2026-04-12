package nl.roeltje.liturgie.services;

import nl.roeltje.liturgie.models.Liturgy;
import nl.roeltje.liturgie.models.LiturgySection;
import nl.roeltje.liturgie.models.LiturgySlide;
import nl.roeltje.liturgie.models.Settings;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.zip.*;

/**
 * Orchestrates export of a liturgy to PPTX, PDF ZIP, links TXT, and Excel.
 */
public class ExportService {

    private static final Logger log = LoggerFactory.getLogger(ExportService.class);

    private Settings settings;
    private final PptxService pptxService;
    private final ExcelService excelService;

    public ExportService(Settings settings, PptxService pptxService, ExcelService excelService) {
        this.settings    = settings;
        this.pptxService = pptxService;
        this.excelService = excelService;
    }

    public void setSettings(Settings settings) { this.settings = settings; }

    // ── Filename helpers ───────────────────────────────────────────────────────

    public String getDefaultFilename(Liturgy liturgy) {
        String pattern = settings.getOutput_pattern();
        String date = liturgy.getService_date() != null
                ? liturgy.getService_date()
                : LocalDate.now().format(DateTimeFormatter.ISO_LOCAL_DATE);
        return pattern
                .replace("{date}", date)
                .replace("{year}",  date.length() >= 4 ? date.substring(0, 4) : "")
                .replace("{month}", date.length() >= 7 ? date.substring(5, 7) : "")
                .replace("{day}",   date.length() >= 10 ? date.substring(8, 10) : "");
    }

    public Path getOutputFolder() throws IOException {
        Path dir = settings.getOutputPath();
        Files.createDirectories(dir);
        return dir;
    }

    // ── PPTX export ────────────────────────────────────────────────────────────

    public Path exportPptx(Liturgy liturgy, Path outputFile) throws IOException {
        Files.createDirectories(outputFile.getParent());
        pptxService.mergeLiturgy(liturgy, outputFile, settings);
        log.info("PPTX exported to {}", outputFile);
        return outputFile;
    }

    // ── PDF ZIP export ─────────────────────────────────────────────────────────

    public Path exportPdfZip(Liturgy liturgy, Path outputFile) throws IOException {
        Files.createDirectories(outputFile.getParent());
        Set<String> seen = new HashSet<>();

        try (ZipOutputStream zos = new ZipOutputStream(
                new BufferedOutputStream(Files.newOutputStream(outputFile,
                        StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING)))) {
            zos.setLevel(Deflater.BEST_COMPRESSION);

            for (LiturgySection section : liturgy.getSections()) {
                String pdfPathStr = section.getPdf_path();
                if (pdfPathStr == null || pdfPathStr.isBlank()) continue;
                Path pdf = Path.of(pdfPathStr);
                if (!Files.exists(pdf)) continue;

                String entryName = sanitizeFilename(section.getName()) + ".pdf";
                // Handle duplicates
                String base = entryName.replace(".pdf", "");
                int counter = 1;
                while (seen.contains(entryName)) {
                    entryName = base + "_" + counter++ + ".pdf";
                }
                seen.add(entryName);

                zos.putNextEntry(new ZipEntry(entryName));
                Files.copy(pdf, zos);
                zos.closeEntry();
            }
        }
        log.info("PDF ZIP exported to {}", outputFile);
        return outputFile;
    }

    // ── Links TXT export ───────────────────────────────────────────────────────

    public Path exportLinksTxt(Liturgy liturgy, Path outputFile) throws IOException {
        Files.createDirectories(outputFile.getParent());
        String content = generateLinksText(liturgy);
        Files.writeString(outputFile, content, java.nio.charset.StandardCharsets.UTF_8);
        log.info("Links TXT exported to {}", outputFile);
        return outputFile;
    }

    public String generateLinksText(Liturgy liturgy) {
        StringBuilder sb = new StringBuilder();
        sb.append(liturgy.getName()).append("\n");
        if (liturgy.getService_date() != null) {
            sb.append(liturgy.getService_date()).append("\n");
        }
        if (liturgy.getDienstleider() != null) {
            sb.append("Dienstleider: ").append(liturgy.getDienstleider()).append("\n");
        }
        sb.append("\n");

        int sectionNum = 1;
        for (LiturgySection section : liturgy.getSections()) {
            sb.append(sectionNum++).append(". ").append(section.getName()).append("\n");
            if (section.isSong()) {
                if (!section.getYoutube_links().isEmpty()) {
                    sb.append("   YouTube:\n");
                    for (String url : section.getYoutube_links()) {
                        sb.append("     ").append(url).append("\n");
                    }
                }
                if (section.getPdf_path() != null) {
                    sb.append("   PDF: ").append(Path.of(section.getPdf_path()).getFileName()).append("\n");
                }
            }
        }
        return sb.toString();
    }

    // ── All formats ────────────────────────────────────────────────────────────

    public record ExportResult(Map<String, Path> files, List<String> errors) {}

    public ExportResult exportAll(Liturgy liturgy,
                                  boolean includePptx,
                                  boolean includePdf,
                                  boolean includeTxt,
                                  boolean includeExcel) {
        Map<String, Path> files = new LinkedHashMap<>();
        List<String> errors = new ArrayList<>();

        try {
            Path outputDir = getOutputFolder();
            String base = stripExtension(getDefaultFilename(liturgy));

            if (includePptx) {
                try {
                    Path out = outputDir.resolve(base + ".pptx");
                    exportPptx(liturgy, out);
                    files.put("pptx", out);
                } catch (IOException e) {
                    errors.add("PPTX: " + e.getMessage());
                    log.error("PPTX export failed", e);
                }
            }

            if (includePdf) {
                try {
                    Path out = outputDir.resolve(base + "_pdfs.zip");
                    exportPdfZip(liturgy, out);
                    files.put("pdf", out);
                } catch (IOException e) {
                    errors.add("PDF ZIP: " + e.getMessage());
                    log.error("PDF ZIP export failed", e);
                }
            }

            if (includeTxt) {
                try {
                    Path out = outputDir.resolve(base + "_links.txt");
                    exportLinksTxt(liturgy, out);
                    files.put("txt", out);
                } catch (IOException e) {
                    errors.add("TXT: " + e.getMessage());
                    log.error("TXT export failed", e);
                }
            }

            if (includeExcel) {
                try {
                    Path xl = settings.getExcelRegisterPath();
                    if (xl != null && Files.exists(xl)) {
                        excelService.exportLiturgy(liturgy, xl);
                        files.put("excel", xl);
                    }
                } catch (IOException e) {
                    errors.add("Excel: " + e.getMessage());
                    log.error("Excel export failed", e);
                }
            }
        } catch (IOException e) {
            errors.add("Output folder: " + e.getMessage());
        }
        return new ExportResult(files, errors);
    }

    // ── Helpers ────────────────────────────────────────────────────────────────

    private String sanitizeFilename(String name) {
        return name.replaceAll("[\\\\/:*?\"<>|]", "_");
    }

    private String stripExtension(String filename) {
        int dot = filename.lastIndexOf('.');
        return dot > 0 ? filename.substring(0, dot) : filename;
    }
}
