package nl.roeltje.liturgie.services;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.net.URI;
import java.net.http.*;
import java.nio.file.*;
import java.time.Duration;
import java.util.*;
import java.util.concurrent.*;
import java.util.regex.*;

/**
 * YouTube service – search and validation via yt-dlp subprocess.
 */
public class YouTubeService {

    private static final Logger log = LoggerFactory.getLogger(YouTubeService.class);
    private static final ObjectMapper JSON = new ObjectMapper();

    private static final Pattern YT_ID_PATTERN = Pattern.compile(
            "(?:youtube\\.com/watch\\?v=|youtu\\.be/|youtube\\.com/embed/)([A-Za-z0-9_-]{11})");

    public record YouTubeResult(String title, String url, String channel,
                                String duration, String thumbnailUrl) {}

    // ── yt-dlp availability ────────────────────────────────────────────────────

    public boolean isYtDlpAvailable() {
        try {
            Process p = new ProcessBuilder("python", "-m", "yt_dlp", "--version")
                    .redirectErrorStream(true).start();
            p.waitFor(5, TimeUnit.SECONDS);
            return p.exitValue() == 0;
        } catch (Exception e) {
            return false;
        }
    }

    // ── Search ─────────────────────────────────────────────────────────────────

    public List<YouTubeResult> search(String query, int maxResults) {
        if (query == null || query.isBlank()) return List.of();
        try {
            ProcessBuilder pb = new ProcessBuilder(
                    "python", "-m", "yt_dlp",
                    "--flat-playlist",
                    "--dump-json",
                    "--no-warnings",
                    "ytsearch" + maxResults + ":" + query);
            pb.redirectErrorStream(false);
            Process proc = pb.start();
            boolean finished = proc.waitFor(30, TimeUnit.SECONDS);
            if (!finished) { proc.destroyForcibly(); return List.of(); }

            String output = new String(proc.getInputStream().readAllBytes());
            List<YouTubeResult> results = new ArrayList<>();
            for (String line : output.split("\n")) {
                line = line.trim();
                if (line.isEmpty()) continue;
                try {
                    JsonNode node = JSON.readTree(line);
                    String title = node.path("title").asText("");
                    String videoId = node.path("id").asText("");
                    String url = videoId.isEmpty() ? "" : "https://www.youtube.com/watch?v=" + videoId;
                    String channel = node.path("uploader").asText(node.path("channel").asText(""));
                    String duration = formatDuration(node.path("duration").asLong(0));
                    String thumb = node.path("thumbnail").asText("");
                    results.add(new YouTubeResult(title, url, channel, duration, thumb));
                } catch (Exception e) {
                    log.debug("Cannot parse yt-dlp line: {}", line);
                }
            }
            return results;
        } catch (IOException | InterruptedException e) {
            log.warn("YouTube search failed: {}", e.getMessage());
            return List.of();
        }
    }

    // ── Validation ─────────────────────────────────────────────────────────────

    public Optional<String> extractVideoId(String url) {
        if (url == null) return Optional.empty();
        Matcher m = YT_ID_PATTERN.matcher(url);
        return m.find() ? Optional.of(m.group(1)) : Optional.empty();
    }

    /** Fast validation via HEAD request. */
    public boolean validateLinkFast(String url) {
        Optional<String> id = extractVideoId(url);
        if (id.isEmpty()) return false;
        try {
            HttpClient client = HttpClient.newBuilder()
                    .connectTimeout(Duration.ofSeconds(10))
                    .followRedirects(HttpClient.Redirect.NORMAL)
                    .build();
            HttpRequest req = HttpRequest.newBuilder()
                    .uri(URI.create("https://www.youtube.com/watch?v=" + id.get()))
                    .method("HEAD", HttpRequest.BodyPublishers.noBody())
                    .timeout(Duration.ofSeconds(10))
                    .build();
            HttpResponse<Void> resp = client.send(req, HttpResponse.BodyHandlers.discarding());
            return resp.statusCode() == 200;
        } catch (Exception e) {
            return false;
        }
    }

    /** Parallel batch validation using virtual threads. */
    public Map<String, Boolean> validateBatch(List<String> urls) {
        Map<String, Boolean> results = new ConcurrentHashMap<>();
        try (ExecutorService pool = Executors.newVirtualThreadPerTaskExecutor()) {
            List<Future<?>> futures = new ArrayList<>();
            for (String url : urls) {
                futures.add(pool.submit(() -> results.put(url, validateLinkFast(url))));
            }
            for (Future<?> f : futures) {
                try { f.get(20, TimeUnit.SECONDS); }
                catch (Exception e) { log.debug("Validation task exception: {}", e.getMessage()); }
            }
        }
        return results;
    }

    // ── File I/O ───────────────────────────────────────────────────────────────

    public List<String> readYoutubeFile(Path songFolder) {
        Path ytFile = songFolder.resolve("youtube.txt");
        if (!Files.exists(ytFile)) return new ArrayList<>();
        try {
            return Files.readAllLines(ytFile).stream()
                    .map(String::trim).filter(l -> !l.isBlank()).toList();
        } catch (IOException e) {
            log.warn("Cannot read {}", ytFile, e);
            return new ArrayList<>();
        }
    }

    public void writeYoutubeFile(Path songFolder, List<String> urls) throws IOException {
        Path ytFile = songFolder.resolve("youtube.txt");
        Files.writeString(ytFile, String.join("\n", urls));
    }

    // ── Helpers ────────────────────────────────────────────────────────────────

    private String formatDuration(long seconds) {
        if (seconds <= 0) return "";
        long m = seconds / 60, s = seconds % 60;
        return String.format("%d:%02d", m, s);
    }
}
