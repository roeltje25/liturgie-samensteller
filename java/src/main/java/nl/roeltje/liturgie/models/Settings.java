package nl.roeltje.liturgie.models;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

/**
 * Persistent application settings.
 *
 * Stored at:
 *   Windows : %APPDATA%\LiturgieSamensteller\settings.json
 *   macOS   : ~/Library/Application Support/LiturgieSamensteller/settings.json
 *   Linux   : ~/.config/LiturgieSamensteller/settings.json
 *
 * Field names are intentionally kept identical to the Python version so that
 * settings files are interchangeable.
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public class Settings {

    private static final Logger log = LoggerFactory.getLogger(Settings.class);
    private static final String APP_NAME = "LiturgieSamensteller";
    private static final ObjectMapper JSON = new ObjectMapper();

    // ── Paths ──────────────────────────────────────────────────────────────────
    private String base_folder = "";
    private String songs_folder = "./Liederen";
    private String algemeen_folder = "./Algemeen";
    private String output_folder = "./Vieringen";
    private String themes_folder = "./Themas";
    private String collecte_filename = "Collecte.pptx";
    private String stub_template_filename = "StubTemplate.pptx";
    private String bible_template_filename = "BijbelTemplate.pptx";
    private String output_pattern = "{date}_viering-generated.pptx";
    private String language = "nl";

    // ── Song cover ─────────────────────────────────────────────────────────────
    private boolean song_cover_enabled = false;
    private String song_cover_filename = "";

    // ── Excel ──────────────────────────────────────────────────────────────────
    private String excel_register_path = "./LiederenRegister.xlsx";
    private String pptx_archive_folder = "./Vieringen";

    // ── Bible slides ──────────────────────────────────────────────────────────
    private String bible_font_name = "Calibri";
    private int bible_font_size = 12;
    private int bible_chars_per_slide = 500;
    private boolean bible_show_verse_numbers = true;
    private String youversion_api_key = "";

    // ── User-curated liturgy item titles (never treated as songs) ─────────────
    private List<String> user_liturgy_items = new ArrayList<>();

    // ── Window state ───────────────────────────────────────────────────────────
    private int window_width = 1200;
    private int window_height = 800;

    // ── Static helpers ─────────────────────────────────────────────────────────

    public static Path getConfigDir() {
        String os = System.getProperty("os.name", "").toLowerCase();
        Path base;
        if (os.contains("win")) {
            String appdata = System.getenv("APPDATA");
            base = appdata != null ? Paths.get(appdata) : Paths.get(System.getProperty("user.home"));
        } else if (os.contains("mac")) {
            base = Paths.get(System.getProperty("user.home"), "Library", "Application Support");
        } else {
            String xdg = System.getenv("XDG_CONFIG_HOME");
            base = xdg != null ? Paths.get(xdg) : Paths.get(System.getProperty("user.home"), ".config");
        }
        return base.resolve(APP_NAME);
    }

    public static Path getSettingsPath() {
        return getConfigDir().resolve("settings.json");
    }

    public static Settings load() {
        Path path = getSettingsPath();
        if (Files.exists(path)) {
            try {
                Settings s = JSON.readValue(path.toFile(), Settings.class);
                log.debug("Settings loaded from {}", path);
                return s;
            } catch (IOException e) {
                log.error("Failed to load settings from {}, using defaults", path, e);
            }
        } else {
            log.debug("Settings file not found, using defaults: {}", path);
        }
        return new Settings();
    }

    public void save() {
        Path path = getSettingsPath();
        try {
            Files.createDirectories(path.getParent());
            JSON.writerWithDefaultPrettyPrinter().writeValue(path.toFile(), this);
            log.debug("Settings saved to {}", path);
        } catch (IOException e) {
            log.error("Failed to save settings to {}", path, e);
        }
    }

    // ── Derived path helpers ────────────────────────────────────────────────────

    public boolean isFirstRun() {
        return base_folder == null || base_folder.isBlank();
    }

    private Path resolveBase() {
        if (base_folder != null && !base_folder.isBlank()) {
            Path p = Paths.get(base_folder);
            if (Files.isDirectory(p)) return p;
        }
        return Paths.get(".");
    }

    private Path resolve(String sub) {
        Path p = Paths.get(sub);
        return p.isAbsolute() ? p : resolveBase().resolve(sub).normalize();
    }

    public Path getSongsPath()          { return resolve(songs_folder); }
    public Path getAlgemeenPath()       { return resolve(algemeen_folder); }
    public Path getOutputPath()         { return resolve(output_folder); }
    public Path getThemesPath()         { return resolve(themes_folder); }
    public Path getCollectePath()       { return getAlgemeenPath().resolve(collecte_filename); }
    public Path getStubTemplatePath()   { return getAlgemeenPath().resolve(stub_template_filename); }
    public Path getBibleTemplatePath()  { return getAlgemeenPath().resolve(bible_template_filename); }
    public Path getPptxArchivePath()    { return resolve(pptx_archive_folder); }

    public Path getSongCoverPath() {
        if (!song_cover_enabled || song_cover_filename.isBlank()) return null;
        Path p = getAlgemeenPath().resolve(song_cover_filename);
        return Files.exists(p) ? p : null;
    }

    public Path getExcelRegisterPath() {
        if (excel_register_path == null || excel_register_path.isBlank()) return null;
        return resolve(excel_register_path);
    }

    // ── Getters / Setters (Jackson uses these) ──────────────────────────────────

    public String getBase_folder() { return base_folder; }
    public void setBase_folder(String v) { base_folder = v; }

    public String getSongs_folder() { return songs_folder; }
    public void setSongs_folder(String v) { songs_folder = v; }

    public String getAlgemeen_folder() { return algemeen_folder; }
    public void setAlgemeen_folder(String v) { algemeen_folder = v; }

    public String getOutput_folder() { return output_folder; }
    public void setOutput_folder(String v) { output_folder = v; }

    public String getThemes_folder() { return themes_folder; }
    public void setThemes_folder(String v) { themes_folder = v; }

    public String getCollecte_filename() { return collecte_filename; }
    public void setCollecte_filename(String v) { collecte_filename = v; }

    public String getStub_template_filename() { return stub_template_filename; }
    public void setStub_template_filename(String v) { stub_template_filename = v; }

    public String getBible_template_filename() { return bible_template_filename; }
    public void setBible_template_filename(String v) { bible_template_filename = v; }

    public String getOutput_pattern() { return output_pattern; }
    public void setOutput_pattern(String v) { output_pattern = v; }

    public String getLanguage() { return language; }
    public void setLanguage(String v) { language = v; }

    public boolean isSong_cover_enabled() { return song_cover_enabled; }
    public void setSong_cover_enabled(boolean v) { song_cover_enabled = v; }

    public String getSong_cover_filename() { return song_cover_filename; }
    public void setSong_cover_filename(String v) { song_cover_filename = v; }

    public String getExcel_register_path() { return excel_register_path; }
    public void setExcel_register_path(String v) { excel_register_path = v; }

    public String getPptx_archive_folder() { return pptx_archive_folder; }
    public void setPptx_archive_folder(String v) { pptx_archive_folder = v; }

    public String getBible_font_name() { return bible_font_name; }
    public void setBible_font_name(String v) { bible_font_name = v; }

    public int getBible_font_size() { return bible_font_size; }
    public void setBible_font_size(int v) { bible_font_size = v; }

    public int getBible_chars_per_slide() { return bible_chars_per_slide; }
    public void setBible_chars_per_slide(int v) { bible_chars_per_slide = v; }

    public boolean isBible_show_verse_numbers() { return bible_show_verse_numbers; }
    public void setBible_show_verse_numbers(boolean v) { bible_show_verse_numbers = v; }

    public String getYouversion_api_key() { return youversion_api_key; }
    public void setYouversion_api_key(String v) { youversion_api_key = v; }

    public List<String> getUser_liturgy_items() { return user_liturgy_items; }
    public void setUser_liturgy_items(List<String> v) { user_liturgy_items = v != null ? v : new ArrayList<>(); }

    public int getWindowWidth() { return window_width; }
    public void setWindowWidth(int v) { window_width = v; }

    public int getWindowHeight() { return window_height; }
    public void setWindowHeight(int v) { window_height = v; }

    // Jackson property names match the Python snake_case fields
    public int getWindow_width() { return window_width; }
    public void setWindow_width(int v) { window_width = v; }

    public int getWindow_height() { return window_height; }
    public void setWindow_height(int v) { window_height = v; }
}
