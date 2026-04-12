package nl.roeltje.liturgie;

import nl.roeltje.liturgie.i18n.TranslationManager;
import nl.roeltje.liturgie.models.Liturgy;
import nl.roeltje.liturgie.models.Settings;
import nl.roeltje.liturgie.services.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Application-level singleton that holds settings, services and the currently
 * open {@link Liturgy}.  Pass this to controllers via setter injection rather
 * than using static access wherever possible.
 */
public final class AppContext {

    public static final String VERSION = "1.2.2";

    private static final Logger log = LoggerFactory.getLogger(AppContext.class);
    private static AppContext instance;

    // ── Core state ─────────────────────────────────────────────────────────
    private Settings settings;
    private Liturgy currentLiturgy;
    private String currentLiturgyPath;   // null = unsaved

    // ── Services ────────────────────────────────────────────────────────────
    private final TranslationManager translations;
    private final FolderScannerService folderScanner;
    private final PptxService pptxService;
    private final ExcelService excelService;
    private final ExportService exportService;
    private final YouTubeService youTubeService;
    private final BibleService bibleService;
    private final BibleSlideService bibleSlideService;
    private final ThemeService themeService;
    private final SongMatcherService songMatcher;

    private AppContext(Settings settings) {
        this.settings = settings;
        this.translations = TranslationManager.getInstance();
        translations.setLanguage(settings.getLanguage());

        this.folderScanner   = new FolderScannerService();
        this.pptxService     = new PptxService();
        this.excelService    = new ExcelService();
        this.exportService   = new ExportService(settings, pptxService, excelService);
        this.youTubeService  = new YouTubeService();
        this.bibleService    = new BibleService(settings);
        this.bibleSlideService = new BibleSlideService(settings, pptxService);
        this.themeService    = new ThemeService();
        this.songMatcher     = new SongMatcherService();
    }

    /** Initialise the singleton (call once from {@link Main#start}). */
    public static AppContext init(Settings settings) {
        instance = new AppContext(settings);
        return instance;
    }

    /** Access the singleton after initialisation. */
    public static AppContext get() {
        if (instance == null) throw new IllegalStateException("AppContext not initialised");
        return instance;
    }

    // ── Translation shorthand ────────────────────────────────────────────────
    public String tr(String key, Object... args) {
        return translations.tr(key, args);
    }

    // ── Settings ─────────────────────────────────────────────────────────────
    public Settings getSettings() { return settings; }

    public void updateSettings(Settings updated) {
        this.settings = updated;
        translations.setLanguage(updated.getLanguage());
        exportService.setSettings(updated);
        bibleService.setSettings(updated);
        bibleSlideService.setSettings(updated);
        updated.save();
        log.debug("Settings updated and saved");
    }

    // ── Current liturgy ───────────────────────────────────────────────────────
    public Liturgy getCurrentLiturgy() { return currentLiturgy; }
    public void setCurrentLiturgy(Liturgy liturgy) { this.currentLiturgy = liturgy; }

    public String getCurrentLiturgyPath() { return currentLiturgyPath; }
    public void setCurrentLiturgyPath(String path) { this.currentLiturgyPath = path; }

    public boolean hasUnsavedLiturgy() { return currentLiturgy != null; }

    // ── Service accessors ─────────────────────────────────────────────────────
    public TranslationManager getTranslations()       { return translations; }
    public FolderScannerService getFolderScanner()    { return folderScanner; }
    public PptxService getPptxService()               { return pptxService; }
    public ExcelService getExcelService()             { return excelService; }
    public ExportService getExportService()           { return exportService; }
    public YouTubeService getYouTubeService()         { return youTubeService; }
    public BibleService getBibleService()             { return bibleService; }
    public BibleSlideService getBibleSlideService()   { return bibleSlideService; }
    public ThemeService getThemeService()             { return themeService; }
    public SongMatcherService getSongMatcher()        { return songMatcher; }
}
