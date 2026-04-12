package nl.roeltje.liturgie.services;

import nl.roeltje.liturgie.models.Settings;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.io.*;
import java.nio.file.*;
import java.util.List;

/**
 * Generates a PPTX file containing Bible verse slides.
 */
public class BibleSlideService {

    private static final Logger log = LoggerFactory.getLogger(BibleSlideService.class);

    private Settings settings;
    private final PptxService pptxService;

    public BibleSlideService(Settings settings, PptxService pptxService) {
        this.settings = settings;
        this.pptxService = pptxService;
    }

    public void setSettings(Settings settings) { this.settings = settings; }

    /**
     * Generate a .pptx with one slide per verse (or verse group fitting within
     * {@code charsPerSlide}) for the given verse list.
     *
     * @return path to generated file
     */
    public Path generateSlides(List<BibleService.BibleVerse> verses,
                               List<String> translationAbbreviations,
                               String reference) throws IOException {
        Path output = Files.createTempFile("bible_", ".pptx");

        // Try to use the Bible template if it exists
        Path template = settings.getBibleTemplatePath();
        XMLSlideShow show;
        if (template != null && Files.exists(template)) {
            show = new XMLSlideShow(Files.newInputStream(template));
            // Clear existing slides from template
            while (!show.getSlides().isEmpty()) {
                show.removeSlide(0);
            }
        } else {
            show = new XMLSlideShow();
        }

        // Group verses into slides based on chars per slide
        int charsPerSlide = settings.getBible_chars_per_slide();
        String fontName = settings.getBible_font_name();
        int fontSize = settings.getBible_font_size();

        // Build text blocks: each block is one slide's worth of content
        List<String> blocks = groupVersesIntoBlocks(verses, charsPerSlide,
                settings.isBible_show_verse_numbers(), translationAbbreviations);

        for (String block : blocks) {
            XSLFSlide slide = show.createSlide();
            addTextToSlide(slide, block, fontName, fontSize);
        }

        try (OutputStream out = Files.newOutputStream(output,
                StandardOpenOption.TRUNCATE_EXISTING)) {
            show.write(out);
        }
        show.close();
        log.info("Bible slides generated: {} slides → {}", blocks.size(), output);
        return output;
    }

    private List<String> groupVersesIntoBlocks(List<BibleService.BibleVerse> verses,
                                               int charsPerSlide,
                                               boolean showVerseNumbers,
                                               List<String> abbrevs) {
        List<String> blocks = new java.util.ArrayList<>();
        StringBuilder current = new StringBuilder();
        for (BibleService.BibleVerse verse : verses) {
            String text = showVerseNumbers
                    ? verse.verse() + " " + verse.text()
                    : verse.text();
            if (current.length() > 0 && current.length() + text.length() > charsPerSlide) {
                blocks.add(current.toString().trim());
                current = new StringBuilder();
            }
            if (!current.isEmpty()) current.append("\n");
            current.append(text);
        }
        if (!current.isEmpty()) blocks.add(current.toString().trim());
        return blocks;
    }

    private void addTextToSlide(XSLFSlide slide, String text, String fontName, int fontSize) {
        java.awt.Dimension pageSize = slide.getSlideShow().getPageSize();
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(new java.awt.geom.Rectangle2D.Double(
                40, 40, pageSize.getWidth() - 80, pageSize.getHeight() - 80));

        XSLFTextParagraph para = box.addNewTextParagraph();
        XSLFTextRun run = para.addNewTextRun();
        run.setText(text);
        run.setFontFamily(fontName);
        run.setFontSize((double) fontSize);
        run.setFontColor(Color.WHITE);
    }
}
