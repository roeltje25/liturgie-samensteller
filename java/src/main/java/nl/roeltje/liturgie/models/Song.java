package nl.roeltje.liturgie.models;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

/**
 * Represents a song in the songs library (a folder containing a .pptx and
 * optional .pdf + youtube.txt).
 */
public class Song {

    private String name;
    private Path folderPath;
    private String relativePath;
    private String title;
    private Path pptxPath;
    private Path pdfPath;
    private List<String> youtubeLinks = new ArrayList<>();

    public Song() {}

    public Song(String name, Path folderPath, String relativePath) {
        this.name = name;
        this.folderPath = folderPath;
        this.relativePath = relativePath;
        this.title = name; // default; overridden by song.properties if present
    }

    // ── Convenience ────────────────────────────────────────────────────────────

    public boolean hasPptx() { return pptxPath != null && Files.exists(pptxPath); }

    public boolean hasPdf() { return pdfPath != null && Files.exists(pdfPath); }

    public boolean hasYoutube() { return !youtubeLinks.isEmpty(); }

    /** Display name shown in the song picker. */
    public String getDisplayTitle() {
        return title != null && !title.isBlank() ? title : name;
    }

    // ── Getters / Setters ───────────────────────────────────────────────────────

    public String getName() { return name; }
    public void setName(String name) { this.name = name; }

    public Path getFolderPath() { return folderPath; }
    public void setFolderPath(Path folderPath) { this.folderPath = folderPath; }

    public String getRelativePath() { return relativePath; }
    public void setRelativePath(String relativePath) { this.relativePath = relativePath; }

    public String getTitle() { return title; }
    public void setTitle(String title) { this.title = title; }

    public Path getPptxPath() { return pptxPath; }
    public void setPptxPath(Path pptxPath) { this.pptxPath = pptxPath; }

    public Path getPdfPath() { return pdfPath; }
    public void setPdfPath(Path pdfPath) { this.pdfPath = pdfPath; }

    public List<String> getYoutubeLinks() { return youtubeLinks; }
    public void setYoutubeLinks(List<String> youtubeLinks) {
        this.youtubeLinks = youtubeLinks != null ? youtubeLinks : new ArrayList<>();
    }

    @Override
    public String toString() {
        return "Song{name='" + name + "', pptx=" + pptxPath + "}";
    }
}
