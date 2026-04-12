package nl.roeltje.liturgie.services;

import nl.roeltje.liturgie.models.Liturgy;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.*;
import java.util.List;
import java.util.stream.Stream;

/**
 * Loads and saves liturgy theme templates (.liturgy files in the themes folder).
 */
public class ThemeService {

    private static final Logger log = LoggerFactory.getLogger(ThemeService.class);

    public List<Path> listThemes(Path themesDir) {
        if (themesDir == null || !Files.isDirectory(themesDir)) return List.of();
        try (Stream<Path> stream = Files.list(themesDir)) {
            return stream
                    .filter(p -> p.getFileName().toString().endsWith(".liturgy"))
                    .sorted()
                    .toList();
        } catch (IOException e) {
            log.warn("Cannot list themes in {}", themesDir, e);
            return List.of();
        }
    }

    public Liturgy loadTheme(Path themeFile) throws IOException {
        Liturgy.MigrationResult result = Liturgy.loadWithMigration(themeFile, themeFile.getParent());
        return result.liturgy();
    }

    public void saveTheme(Liturgy liturgy, Path themeFile) throws IOException {
        liturgy.save(themeFile, themeFile.getParent());
        log.info("Theme saved to {}", themeFile);
    }
}
