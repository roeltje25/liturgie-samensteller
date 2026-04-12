package nl.roeltje.liturgie.models;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonInclude;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

/**
 * A single slide reference within a {@link LiturgySection}.
 * Field names match the Python v2 model for JSON compatibility.
 */
@JsonIgnoreProperties(ignoreUnknown = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class LiturgySlide {

    private String id = UUID.randomUUID().toString();
    private String title = "";
    private int slide_index = 0;
    private String source_path;
    private Map<String, String> fields = new HashMap<>();
    private boolean is_stub = false;
    private String pdf_path;
    private List<String> youtube_links = new ArrayList<>();

    public LiturgySlide() {}

    /** Deep copy with a new UUID. */
    public LiturgySlide copy() {
        LiturgySlide c = new LiturgySlide();
        c.id = UUID.randomUUID().toString();
        c.title = title;
        c.slide_index = slide_index;
        c.source_path = source_path;
        c.fields = new HashMap<>(fields);
        c.is_stub = is_stub;
        c.pdf_path = pdf_path;
        c.youtube_links = new ArrayList<>(youtube_links);
        return c;
    }

    // ── Getters / Setters ───────────────────────────────────────────────────

    public String getId() { return id; }
    public void setId(String id) { this.id = id; }

    public String getTitle() { return title; }
    public void setTitle(String title) { this.title = title != null ? title : ""; }

    public int getSlide_index() { return slide_index; }
    public void setSlide_index(int slide_index) { this.slide_index = slide_index; }

    public String getSource_path() { return source_path; }
    public void setSource_path(String source_path) { this.source_path = source_path; }

    public Map<String, String> getFields() { return fields; }
    public void setFields(Map<String, String> fields) { this.fields = fields != null ? fields : new HashMap<>(); }

    public boolean isIs_stub() { return is_stub; }
    public void setIs_stub(boolean is_stub) { this.is_stub = is_stub; }

    public String getPdf_path() { return pdf_path; }
    public void setPdf_path(String pdf_path) { this.pdf_path = pdf_path; }

    public List<String> getYoutube_links() { return youtube_links; }
    public void setYoutube_links(List<String> youtube_links) {
        this.youtube_links = youtube_links != null ? youtube_links : new ArrayList<>();
    }

    @Override
    public String toString() {
        return "LiturgySlide{id=" + id + ", title='" + title + "', source=" + source_path + "}";
    }
}
