package nl.roeltje.liturgie.services;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import nl.roeltje.liturgie.models.Settings;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.net.URI;
import java.net.http.*;
import java.time.Duration;
import java.util.*;

/**
 * Fetches Bible verses from the YouVersion API.
 */
public class BibleService {

    private static final Logger log = LoggerFactory.getLogger(BibleService.class);
    private static final ObjectMapper JSON = new ObjectMapper();
    private static final String API_BASE = "https://developers.youversion.com/api/v1";

    private Settings settings;
    private final HttpClient http = HttpClient.newBuilder()
            .connectTimeout(Duration.ofSeconds(15))
            .followRedirects(HttpClient.Redirect.NORMAL)
            .build();

    public record BibleVerse(int chapter, int verse, String text) {}
    public record BibleTranslation(String id, String abbreviation, String name, String language) {}

    public BibleService(Settings settings) {
        this.settings = settings;
    }

    public void setSettings(Settings settings) { this.settings = settings; }

    // ── Translations list ──────────────────────────────────────────────────────

    public List<BibleTranslation> getTranslations(String languageFilter) {
        String apiKey = settings.getYouversion_api_key();
        if (apiKey == null || apiKey.isBlank()) return List.of();
        try {
            String url = API_BASE + "/versions?language_tag=" +
                    (languageFilter != null ? languageFilter : "");
            JsonNode root = get(url, apiKey);
            List<BibleTranslation> result = new ArrayList<>();
            if (root.isArray()) {
                for (JsonNode n : root) {
                    result.add(new BibleTranslation(
                            n.path("id").asText(),
                            n.path("abbreviation").asText(),
                            n.path("local_title").asText(n.path("title").asText()),
                            n.path("language_tag").asText()));
                }
            } else {
                for (JsonNode n : root.path("data")) {
                    result.add(new BibleTranslation(
                            n.path("id").asText(),
                            n.path("abbreviation").asText(),
                            n.path("local_title").asText(n.path("title").asText()),
                            n.path("language_tag").asText()));
                }
            }
            return result;
        } catch (Exception e) {
            log.warn("Cannot fetch translations: {}", e.getMessage());
            return List.of();
        }
    }

    // ── Verse fetching ─────────────────────────────────────────────────────────

    public List<BibleVerse> fetchVerses(String reference, String versionId) {
        String apiKey = settings.getYouversion_api_key();
        if (apiKey == null || apiKey.isBlank()) {
            log.warn("No YouVersion API key configured");
            return List.of();
        }
        try {
            String encoded = java.net.URLEncoder.encode(reference, java.nio.charset.StandardCharsets.UTF_8);
            String url = API_BASE + "/passages?passage=" + encoded + "&version_id=" + versionId;
            JsonNode root = get(url, apiKey);
            List<BibleVerse> verses = new ArrayList<>();
            for (JsonNode verse : root.path("verses")) {
                verses.add(new BibleVerse(
                        verse.path("chapter_number").asInt(),
                        verse.path("verse_start").asInt(),
                        verse.path("content").asText()));
            }
            return verses;
        } catch (Exception e) {
            log.warn("Cannot fetch verses for '{}': {}", reference, e.getMessage());
            return List.of();
        }
    }

    // ── Helpers ────────────────────────────────────────────────────────────────

    private JsonNode get(String url, String apiKey) throws Exception {
        HttpRequest req = HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("X-YouVersion-Developer-Token", apiKey)
                .GET()
                .timeout(Duration.ofSeconds(20))
                .build();
        HttpResponse<String> resp = http.send(req, HttpResponse.BodyHandlers.ofString());
        if (resp.statusCode() != 200) {
            throw new RuntimeException("HTTP " + resp.statusCode() + " from YouVersion API");
        }
        return JSON.readTree(resp.body());
    }
}
