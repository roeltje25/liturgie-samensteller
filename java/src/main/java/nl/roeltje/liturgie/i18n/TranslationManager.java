package nl.roeltje.liturgie.i18n;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CopyOnWriteArrayList;

/**
 * Singleton translation manager.
 *
 * Loads translations from classpath resources {@code i18n/nl.json} and
 * {@code i18n/en.json}.  Keys use dot-path notation (e.g.
 * {@code "menu.file.new"}).  Placeholders follow Python str.format() style:
 * {@code {name}}, {@code {count}}, etc.
 */
public final class TranslationManager {

    private static final Logger log = LoggerFactory.getLogger(TranslationManager.class);
    private static final ObjectMapper JSON = new ObjectMapper();
    private static final TranslationManager INSTANCE = new TranslationManager();

    private Map<String, String> translations;
    private String currentLanguage = "nl";
    private final List<Runnable> listeners = new CopyOnWriteArrayList<>();

    private TranslationManager() {
        load("nl");
    }

    public static TranslationManager getInstance() {
        return INSTANCE;
    }

    /** Switch the active language and notify all registered listeners. */
    public void setLanguage(String lang) {
        if (lang == null || lang.isBlank()) return;
        if (lang.equals(currentLanguage) && translations != null) return;
        load(lang);
        for (Runnable l : listeners) {
            try { l.run(); } catch (Exception e) { log.warn("Language listener threw", e); }
        }
    }

    public String getLanguage() { return currentLanguage; }

    /**
     * Look up a translation key and optionally interpolate {@code {name}}-style
     * placeholders.
     *
     * Example: {@code tr("dialog.song.stub_selected", "title", "Amazing Grace")}
     *
     * Args are passed as alternating key/value pairs:
     *   tr(key, "name", value, "count", 3)
     */
    public String tr(String key, Object... args) {
        String raw = translations != null ? translations.getOrDefault(key, key) : key;
        if (args.length == 0) return raw;

        // Replace {placeholder} patterns with provided values
        for (int i = 0; i + 1 < args.length; i += 2) {
            String placeholder = "{" + args[i] + "}";
            String value = String.valueOf(args[i + 1]);
            raw = raw.replace(placeholder, value);
        }
        return raw;
    }

    /** Register a listener to be called when the language changes. */
    public void addLanguageChangeListener(Runnable listener) {
        listeners.add(listener);
    }

    public void removeLanguageChangeListener(Runnable listener) {
        listeners.remove(listener);
    }

    // ── Private ────────────────────────────────────────────────────────────────

    private void load(String lang) {
        String resource = "/i18n/" + lang + ".json";
        try (InputStream in = TranslationManager.class.getResourceAsStream(resource)) {
            if (in == null) {
                log.warn("Translation resource not found: {}", resource);
                if (translations == null) translations = Map.of();
                return;
            }
            translations = JSON.readValue(in, new TypeReference<Map<String, String>>() {});
            currentLanguage = lang;
            log.debug("Loaded {} translations for language '{}'", translations.size(), lang);
        } catch (Exception e) {
            log.error("Failed to load translations from {}", resource, e);
            if (translations == null) translations = Map.of();
        }
    }
}
