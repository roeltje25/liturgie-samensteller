package nl.roeltje.liturgie.models;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/**
 * A complete liturgy.  Supports both the v2 format (sections/slides) and
 * reading v1 files (elements list), which are automatically migrated on load.
 *
 * Field names are kept identical to the Python model so that .liturgy JSON
 * files remain interchangeable between the Python and Java apps.
 */
@JsonIgnoreProperties(ignoreUnknown = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class Liturgy {

    private static final Logger log = LoggerFactory.getLogger(Liturgy.class);
    private static final ObjectMapper JSON = new ObjectMapper();

    private int format_version = 2;
    private String name = "";
    private String created_date = LocalDate.now().toString();
    private String theme_source_path;
    private List<LiturgySection> sections = new ArrayList<>();
    private String service_date;
    private String dienstleider;

    public Liturgy() {}

    public Liturgy(String name) {
        this.name = name;
    }

    // ── Section helpers ─────────────────────────────────────────────────────────

    public void addSection(LiturgySection section) {
        sections.add(section);
    }

    public void insertSection(int index, LiturgySection section) {
        int i = Math.max(0, Math.min(index, sections.size()));
        sections.add(i, section);
    }

    public void removeSection(int index) {
        if (index >= 0 && index < sections.size()) sections.remove(index);
    }

    public void moveSection(int from, int to) {
        if (from < 0 || from >= sections.size() || to < 0 || to >= sections.size()) return;
        LiturgySection s = sections.remove(from);
        sections.add(to, s);
    }

    public void moveSlideWithinSection(int sectionIdx, int fromSlide, int toSlide) {
        if (sectionIdx < 0 || sectionIdx >= sections.size()) return;
        List<LiturgySlide> slides = sections.get(sectionIdx).getSlides();
        if (fromSlide < 0 || fromSlide >= slides.size() || toSlide < 0 || toSlide >= slides.size()) return;
        LiturgySlide slide = slides.remove(fromSlide);
        slides.add(toSlide, slide);
    }

    public Optional<LiturgySection> findSectionById(String id) {
        return sections.stream().filter(s -> s.getId().equals(id)).findFirst();
    }

    public Optional<LiturgySlide> findSlideById(String id) {
        return sections.stream()
                .flatMap(s -> s.getSlides().stream())
                .filter(sl -> sl.getId().equals(id))
                .findFirst();
    }

    // ── Persistence ─────────────────────────────────────────────────────────────

    /**
     * Save to a .liturgy JSON file.
     *
     * @param filePath  absolute path to write
     * @param basePath  if non-null, absolute paths are stored relative to this
     */
    public void save(Path filePath, Path basePath) throws IOException {
        Liturgy toWrite = basePath != null ? withRelativePaths(basePath) : this;
        Files.createDirectories(filePath.getParent());
        JSON.writerWithDefaultPrettyPrinter().writeValue(filePath.toFile(), toWrite);
        log.info("Liturgy saved to {}", filePath);
    }

    /**
     * Load a .liturgy file, automatically migrating v1 format if needed.
     *
     * @return (liturgy, wasMigrated)
     */
    public static MigrationResult loadWithMigration(Path filePath, Path basePath) throws IOException {
        Liturgy liturgy = JSON.readValue(filePath.toFile(), Liturgy.class);
        boolean migrated = false;

        if (liturgy.format_version < 2) {
            liturgy = migrateV1(liturgy);
            migrated = true;
        }

        if (basePath != null) {
            liturgy = liturgy.withAbsolutePaths(basePath);
        }
        log.info("Liturgy loaded from {} (migrated={})", filePath, migrated);
        return new MigrationResult(liturgy, migrated);
    }

    public record MigrationResult(Liturgy liturgy, boolean wasMigrated) {}

    // ── Path conversion ──────────────────────────────────────────────────────────

    /** Return a copy with all absolute paths made relative to {@code base}. */
    private Liturgy withRelativePaths(Path base) {
        Liturgy copy = shallowCopy();
        copy.sections = new ArrayList<>();
        for (LiturgySection sec : sections) {
            LiturgySection cs = new LiturgySection();
            cs.setId(sec.getId());
            cs.setName(sec.getName());
            cs.setSection_type(sec.getSection_type());
            cs.setSource_theme_path(toRelative(sec.getSource_theme_path(), base));
            cs.setPdf_path(toRelative(sec.getPdf_path(), base));
            cs.setYoutube_links(sec.getYoutube_links());
            cs.setSong_source_path(toRelative(sec.getSong_source_path(), base));
            for (LiturgySlide sl : sec.getSlides()) {
                LiturgySlide csl = sl.copy();
                csl.setSource_path(toRelative(sl.getSource_path(), base));
                csl.setPdf_path(toRelative(sl.getPdf_path(), base));
                cs.getSlides().add(csl);
            }
            copy.sections.add(cs);
        }
        return copy;
    }

    /** Return a copy with all relative paths resolved against {@code base}. */
    private Liturgy withAbsolutePaths(Path base) {
        Liturgy copy = shallowCopy();
        copy.sections = new ArrayList<>();
        for (LiturgySection sec : sections) {
            LiturgySection cs = new LiturgySection();
            cs.setId(sec.getId());
            cs.setName(sec.getName());
            cs.setSection_type(sec.getSection_type());
            cs.setSource_theme_path(toAbsolute(sec.getSource_theme_path(), base));
            cs.setPdf_path(toAbsolute(sec.getPdf_path(), base));
            cs.setYoutube_links(sec.getYoutube_links());
            cs.setSong_source_path(toAbsolute(sec.getSong_source_path(), base));
            for (LiturgySlide sl : sec.getSlides()) {
                LiturgySlide csl = sl.copy();
                csl.setSource_path(toAbsolute(sl.getSource_path(), base));
                csl.setPdf_path(toAbsolute(sl.getPdf_path(), base));
                cs.getSlides().add(csl);
            }
            copy.sections.add(cs);
        }
        return copy;
    }

    private static String toRelative(String p, Path base) {
        if (p == null || p.isBlank()) return p;
        Path abs = Paths.get(p);
        if (!abs.isAbsolute()) return p;
        try {
            return base.relativize(abs).toString().replace('\\', '/');
        } catch (IllegalArgumentException e) {
            return p; // different drives
        }
    }

    private static String toAbsolute(String p, Path base) {
        if (p == null || p.isBlank()) return p;
        Path path = Paths.get(p);
        if (path.isAbsolute()) return p;
        return base.resolve(path).normalize().toString();
    }

    private Liturgy shallowCopy() {
        Liturgy c = new Liturgy();
        c.format_version = 2;
        c.name = name;
        c.created_date = created_date;
        c.theme_source_path = theme_source_path;
        c.service_date = service_date;
        c.dienstleider = dienstleider;
        return c;
    }

    // ── V1 migration ─────────────────────────────────────────────────────────────

    /**
     * V1 files have an "elements" array instead of "sections".
     * Jackson will put them into _v1Elements if annotated; here we just
     * handle the conversion after the fact.
     */
    private List<Object> _v1_elements; // raw, set by Jackson for v1 files

    @com.fasterxml.jackson.annotation.JsonProperty("elements")
    public void setElements(List<Object> elements) {
        this._v1_elements = elements;
    }

    @com.fasterxml.jackson.annotation.JsonIgnore
    public List<Object> getElements_internal() { return _v1_elements; }

    private static Liturgy migrateV1(Liturgy v1) {
        Liturgy v2 = new Liturgy(v1.name);
        v2.created_date = v1.created_date;
        v2.service_date = v1.service_date;
        v2.dienstleider = v1.dienstleider;
        // V1 elements are generic maps from Jackson; create stub sections
        if (v1._v1_elements != null) {
            for (Object raw : v1._v1_elements) {
                if (raw instanceof java.util.Map<?,?> map) {
                    LiturgySection sec = sectionFromV1Map(map);
                    v2.sections.add(sec);
                }
            }
        }
        log.info("Migrated v1 liturgy '{}' to v2 ({} sections)", v1.name, v2.sections.size());
        return v2;
    }

    @SuppressWarnings("unchecked")
    private static LiturgySection sectionFromV1Map(java.util.Map<?,?> m) {
        String type = (String) m.getOrDefault("type", "generic");
        String title = (String) m.getOrDefault("title", "");
        String pptxPath = (String) m.get("pptx_path");
        String pdfPath  = (String) m.get("pdf_path");
        List<String> ytLinks = (List<String>) m.getOrDefault("youtube_links", new ArrayList<>());
        boolean isStub = Boolean.TRUE.equals(m.get("is_stub"));
        Object slideIdxObj = m.get("slide_index");
        int slideIndex = slideIdxObj instanceof Number n ? n.intValue() : 0;

        LiturgySection sec = new LiturgySection();
        sec.setName(title);

        LiturgySlide slide = new LiturgySlide();
        slide.setTitle(title);
        slide.setSlide_index(slideIndex);
        slide.setIs_stub(isStub);

        if ("song".equals(type)) {
            sec.setSection_type(SectionType.SONG);
            sec.setPdf_path(pdfPath);
            sec.setYoutube_links(ytLinks);
            slide.setSource_path(pptxPath);
            slide.setPdf_path(pdfPath);
            slide.setYoutube_links(ytLinks);
        } else {
            sec.setSection_type(SectionType.REGULAR);
            slide.setSource_path(pptxPath != null ? pptxPath : (String) m.get("source_path"));
        }
        sec.getSlides().add(slide);
        return sec;
    }

    // ── Getters / Setters ───────────────────────────────────────────────────────

    public int getFormat_version() { return format_version; }
    public void setFormat_version(int format_version) { this.format_version = format_version; }

    public String getName() { return name; }
    public void setName(String name) { this.name = name != null ? name : ""; }

    public String getCreated_date() { return created_date; }
    public void setCreated_date(String created_date) { this.created_date = created_date; }

    public String getTheme_source_path() { return theme_source_path; }
    public void setTheme_source_path(String theme_source_path) { this.theme_source_path = theme_source_path; }

    public List<LiturgySection> getSections() { return sections; }
    public void setSections(List<LiturgySection> sections) { this.sections = sections != null ? sections : new ArrayList<>(); }

    public String getService_date() { return service_date; }
    public void setService_date(String service_date) { this.service_date = service_date; }

    public String getDienstleider() { return dienstleider; }
    public void setDienstleider(String dienstleider) { this.dienstleider = dienstleider; }

    @Override
    public String toString() {
        return "Liturgy{name='" + name + "', sections=" + sections.size() + "}";
    }
}
