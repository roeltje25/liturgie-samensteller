package nl.roeltje.liturgie.models;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonInclude;

import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

/**
 * A section in the liturgy, containing one or more {@link LiturgySlide}s.
 * Field names match the Python v2 model for JSON compatibility.
 */
@JsonIgnoreProperties(ignoreUnknown = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class LiturgySection {

    private String id = UUID.randomUUID().toString();
    private String name = "";
    private SectionType section_type = SectionType.REGULAR;
    private String source_theme_path;
    private List<LiturgySlide> slides = new ArrayList<>();

    // Song-specific (only meaningful when section_type == SONG)
    private String pdf_path;
    private List<String> youtube_links = new ArrayList<>();
    private String song_source_path;

    public LiturgySection() {}

    /** Deep copy with new UUIDs for section and all slides. */
    public LiturgySection copy() {
        LiturgySection c = new LiturgySection();
        c.id = UUID.randomUUID().toString();
        c.name = name;
        c.section_type = section_type;
        c.source_theme_path = source_theme_path;
        c.pdf_path = pdf_path;
        c.youtube_links = new ArrayList<>(youtube_links);
        c.song_source_path = song_source_path;
        for (LiturgySlide slide : slides) {
            c.slides.add(slide.copy());
        }
        return c;
    }

    // ── Convenience ────────────────────────────────────────────────────────────

    public boolean isSong() { return section_type == SectionType.SONG; }

    public boolean hasYoutube() { return youtube_links != null && !youtube_links.isEmpty(); }

    public boolean hasPdf() { return pdf_path != null && !pdf_path.isBlank(); }

    public boolean hasPptx() {
        return slides.stream().anyMatch(s -> s.getSource_path() != null);
    }

    // ── Getters / Setters ───────────────────────────────────────────────────────

    public String getId() { return id; }
    public void setId(String id) { this.id = id; }

    public String getName() { return name; }
    public void setName(String name) { this.name = name != null ? name : ""; }

    public SectionType getSection_type() { return section_type; }
    public void setSection_type(SectionType section_type) {
        this.section_type = section_type != null ? section_type : SectionType.REGULAR;
    }

    public String getSource_theme_path() { return source_theme_path; }
    public void setSource_theme_path(String source_theme_path) { this.source_theme_path = source_theme_path; }

    public List<LiturgySlide> getSlides() { return slides; }
    public void setSlides(List<LiturgySlide> slides) { this.slides = slides != null ? slides : new ArrayList<>(); }

    public String getPdf_path() { return pdf_path; }
    public void setPdf_path(String pdf_path) { this.pdf_path = pdf_path; }

    public List<String> getYoutube_links() { return youtube_links; }
    public void setYoutube_links(List<String> youtube_links) {
        this.youtube_links = youtube_links != null ? youtube_links : new ArrayList<>();
    }

    public String getSong_source_path() { return song_source_path; }
    public void setSong_source_path(String song_source_path) { this.song_source_path = song_source_path; }

    @Override
    public String toString() {
        return "LiturgySection{id=" + id + ", name='" + name + "', type=" + section_type + ", slides=" + slides.size() + "}";
    }
}
