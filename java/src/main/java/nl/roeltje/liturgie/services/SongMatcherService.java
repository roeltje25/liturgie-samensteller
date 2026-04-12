package nl.roeltje.liturgie.services;

import nl.roeltje.liturgie.models.Song;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.*;

/**
 * Fuzzy song title matching for the QuickLiturgy dialog.
 * Uses Levenshtein distance + token overlap for ranking.
 */
public class SongMatcherService {

    private static final Logger log = LoggerFactory.getLogger(SongMatcherService.class);

    public record MatchResult(String input, Song matched, double score) {}

    /**
     * Find the best match for each input line from the song library.
     *
     * @param inputs   list of song title strings typed by the user
     * @param songs    song library to search
     * @return list of match results, one per input
     */
    public List<MatchResult> matchAll(List<String> inputs, List<Song> songs) {
        List<MatchResult> results = new ArrayList<>();
        for (String input : inputs) {
            if (input == null || input.isBlank()) continue;
            MatchResult best = findBest(input.trim(), songs);
            results.add(best);
        }
        return results;
    }

    private MatchResult findBest(String input, List<Song> songs) {
        Song bestSong = null;
        double bestScore = -1;
        String normalInput = normalize(input);

        for (Song song : songs) {
            double score = similarity(normalInput, normalize(song.getDisplayTitle()));
            if (score > bestScore) {
                bestScore = score;
                bestSong = song;
            }
        }
        return new MatchResult(input, bestSong, bestScore);
    }

    /** Combined similarity: token overlap + normalized edit distance. */
    private double similarity(String a, String b) {
        if (a.equals(b)) return 1.0;
        double tokenScore = tokenOverlap(a, b);
        double editScore = 1.0 - (double) levenshtein(a, b) / Math.max(a.length(), b.length());
        return 0.6 * tokenScore + 0.4 * editScore;
    }

    private double tokenOverlap(String a, String b) {
        Set<String> tokensA = new HashSet<>(Arrays.asList(a.split("\\s+")));
        Set<String> tokensB = new HashSet<>(Arrays.asList(b.split("\\s+")));
        long common = tokensA.stream().filter(tokensB::contains).count();
        int total = tokensA.size() + tokensB.size();
        return total == 0 ? 0 : 2.0 * common / total;
    }

    private int levenshtein(String a, String b) {
        int m = a.length(), n = b.length();
        int[][] dp = new int[m + 1][n + 1];
        for (int i = 0; i <= m; i++) dp[i][0] = i;
        for (int j = 0; j <= n; j++) dp[0][j] = j;
        for (int i = 1; i <= m; i++) {
            for (int j = 1; j <= n; j++) {
                dp[i][j] = a.charAt(i - 1) == b.charAt(j - 1)
                        ? dp[i - 1][j - 1]
                        : 1 + Math.min(dp[i - 1][j - 1], Math.min(dp[i - 1][j], dp[i][j - 1]));
            }
        }
        return dp[m][n];
    }

    private String normalize(String s) {
        return s.toLowerCase()
                .replaceAll("[^a-z0-9 ]", " ")
                .replaceAll("\\s+", " ")
                .trim();
    }
}
