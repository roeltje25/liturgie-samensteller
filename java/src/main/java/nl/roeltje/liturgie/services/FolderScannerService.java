package nl.roeltje.liturgie.services;

import nl.roeltje.liturgie.models.Song;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Optional;
import java.util.stream.Stream;

/**
 * Scans the folder structure for songs, generic items and offering slides.
 * Results are cached; call {@link #refresh()} to invalidate caches.
 */
public class FolderScannerService {

    private static final Logger log = LoggerFactory.getLogger(FolderScannerService.class);

    private List<Song> songsCache;
    private List<Path> genericCache;
    private List<OfferingSlide> offeringsCache;

    public record OfferingSlide(String title, int slideIndex, Path sourcePath) {}

    // ── Songs ──────────────────────────────────────────────────────────────────

    public List<Song> scanSongs(Path songsDir) {
        if (songsCache != null) return songsCache;
        if (songsDir == null || !Files.isDirectory(songsDir)) {
            log.warn("Songs directory not found: {}", songsDir);
            return List.of();
        }
        List<Song> result = new ArrayList<>();
        scanSongsRecursive(songsDir, songsDir, result);
        songsCache = Collections.unmodifiableList(result);
        log.debug("Scanned {} songs from {}", result.size(), songsDir);
        return songsCache;
    }

    private void scanSongsRecursive(Path root, Path dir, List<Song> out) {
        List<Path> children;
        try (Stream<Path> stream = Files.list(dir)) {
            children = stream.sorted().toList();
        } catch (IOException e) {
            log.warn("Cannot list {}", dir, e);
            return;
        }

        boolean hasSongFile = children.stream().anyMatch(this::isSongFile);

        if (hasSongFile) {
            // This is a leaf song folder
            Song song = buildSong(root, dir, children);
            out.add(song);
        } else {
            // Recurse into subdirectories
            for (Path child : children) {
                if (Files.isDirectory(child)) {
                    scanSongsRecursive(root, child, out);
                }
            }
        }
    }

    private boolean isSongFile(Path p) {
        String n = p.getFileName().toString().toLowerCase();
        return n.endsWith(".pptx") || n.endsWith(".ppt") || n.endsWith(".pdf");
    }

    private Song buildSong(Path root, Path dir, List<Path> children) {
        String name = dir.getFileName().toString();
        String rel;
        try {
            rel = root.relativize(dir).toString().replace('\\', '/');
        } catch (IllegalArgumentException e) {
            rel = name;
        }

        Song song = new Song(name, dir, rel);

        // Title from song.properties
        Path props = dir.resolve("song.properties");
        if (Files.exists(props)) {
            try {
                java.util.Properties p = new java.util.Properties();
                try (InputStream in = Files.newInputStream(props)) { p.load(in); }
                String t = p.getProperty("title");
                if (t != null && !t.isBlank()) song.setTitle(t.trim());
            } catch (IOException e) {
                log.debug("Cannot read song.properties in {}", dir);
            }
        }

        // PPTX
        children.stream()
                .filter(p -> p.getFileName().toString().toLowerCase().endsWith(".pptx"))
                .findFirst()
                .ifPresent(song::setPptxPath);

        // PDF
        children.stream()
                .filter(p -> p.getFileName().toString().toLowerCase().endsWith(".pdf"))
                .findFirst()
                .ifPresent(song::setPdfPath);

        // youtube.txt
        Path ytFile = dir.resolve("youtube.txt");
        if (Files.exists(ytFile)) {
            try {
                List<String> lines = Files.readAllLines(ytFile)
                        .stream().map(String::trim)
                        .filter(l -> !l.isBlank()).toList();
                song.setYoutubeLinks(new ArrayList<>(lines));
            } catch (IOException e) {
                log.debug("Cannot read youtube.txt in {}", dir);
            }
        }

        return song;
    }

    // ── Generic (Algemeen) ─────────────────────────────────────────────────────

    private static final List<String> EXCLUDED_ALGEMEEN = List.of(
            "collecte", "stub", "bijbel", "bible");

    public List<Path> scanGeneral(Path algemeenDir) {
        if (genericCache != null) return genericCache;
        if (algemeenDir == null || !Files.isDirectory(algemeenDir)) {
            log.warn("Algemeen directory not found: {}", algemeenDir);
            return List.of();
        }
        List<Path> result;
        try (Stream<Path> stream = Files.list(algemeenDir)) {
            result = stream
                    .filter(p -> !Files.isDirectory(p))
                    .filter(p -> p.getFileName().toString().toLowerCase().endsWith(".pptx"))
                    .filter(p -> {
                        String lower = p.getFileName().toString().toLowerCase();
                        return EXCLUDED_ALGEMEEN.stream().noneMatch(lower::startsWith);
                    })
                    .sorted()
                    .toList();
        } catch (IOException e) {
            log.warn("Cannot scan algemeen dir {}", algemeenDir, e);
            result = List.of();
        }
        genericCache = Collections.unmodifiableList(result);
        log.debug("Scanned {} generic items from {}", result.size(), algemeenDir);
        return genericCache;
    }

    // ── Offering slides ────────────────────────────────────────────────────────

    public List<OfferingSlide> getOfferingSlides(Path collectePptx) {
        if (collectePptx == null || !Files.exists(collectePptx)) {
            log.warn("Collecte PPTX not found: {}", collectePptx);
            return List.of();
        }
        // Cache only for default path; custom paths always re-read
        if (offeringsCache != null) return offeringsCache;

        List<OfferingSlide> result = new ArrayList<>();
        try (XMLSlideShow pptx = new XMLSlideShow(Files.newInputStream(collectePptx))) {
            var slides = pptx.getSlides();
            for (int i = 0; i < slides.size(); i++) {
                String title = extractSlideTitle(slides.get(i), i);
                result.add(new OfferingSlide(title, i, collectePptx));
            }
        } catch (IOException e) {
            log.error("Failed to read collecte PPTX {}", collectePptx, e);
        }
        offeringsCache = Collections.unmodifiableList(result);
        return offeringsCache;
    }

    private String extractSlideTitle(org.apache.poi.xslf.usermodel.XSLFSlide slide, int index) {
        // Try title shape first
        if (slide.getTitle() != null && !slide.getTitle().isBlank()) {
            return slide.getTitle();
        }
        // Fall back to first text shape
        for (var shape : slide.getShapes()) {
            if (shape instanceof org.apache.poi.xslf.usermodel.XSLFTextShape ts) {
                String t = ts.getText();
                if (t != null && !t.isBlank()) return t.trim();
            }
        }
        return "Dia " + (index + 1);
    }

    // ── Lookup ─────────────────────────────────────────────────────────────────

    public Optional<Song> findSongByRelativePath(List<Song> songs, String relativePath) {
        if (relativePath == null) return Optional.empty();
        return songs.stream()
                .filter(s -> relativePath.equals(s.getRelativePath()))
                .findFirst();
    }

    // ── Cache management ───────────────────────────────────────────────────────

    public void refresh() {
        songsCache = null;
        genericCache = null;
        offeringsCache = null;
        log.debug("Folder scanner cache cleared");
    }
}
