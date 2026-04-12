package nl.roeltje.liturgie;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.stage.Stage;
import nl.roeltje.liturgie.models.Settings;
import nl.roeltje.liturgie.ui.MainController;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import atlantafx.base.theme.PrimerLight;

import java.io.IOException;
import java.net.URL;

/**
 * Application entry point.
 *
 * Start-up sequence:
 * 1. Load settings from AppData
 * 2. Apply AtlantaFX theme
 * 3. Show first-run folder picker if not configured
 * 4. Open the main window
 */
public class Main extends Application {

    private static final Logger log = LoggerFactory.getLogger(Main.class);

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        // Apply AtlantaFX PrimerLight theme before any UI is shown
        Application.setUserAgentStylesheet(new PrimerLight().getUserAgentStylesheet());

        // Load (or create default) settings
        Settings settings = Settings.load();
        AppContext context = AppContext.init(settings);

        log.info("Liturgie Samensteller starting, version {}", AppContext.VERSION);

        try {
            URL fxml = getClass().getResource("/nl/roeltje/liturgie/fxml/main-window.fxml");
            if (fxml == null) {
                throw new IOException("FXML not found: main-window.fxml");
            }
            FXMLLoader loader = new FXMLLoader(fxml);
            Scene scene = new Scene(loader.load(),
                    settings.getWindowWidth(),
                    settings.getWindowHeight());

            // Apply app-specific CSS overrides
            URL css = getClass().getResource("/nl/roeltje/liturgie/css/app.css");
            if (css != null) {
                scene.getStylesheets().add(css.toExternalForm());
            }

            MainController controller = loader.getController();
            controller.initStage(primaryStage);

            primaryStage.setTitle(context.tr("app.title"));
            primaryStage.setMinWidth(900);
            primaryStage.setMinHeight(600);
            primaryStage.setScene(scene);
            primaryStage.show();

            // First-run: ask user to pick base folder
            if (settings.isFirstRun()) {
                Platform.runLater(controller::showFirstRunDialog);
            }
        } catch (IOException e) {
            log.error("Failed to load main window", e);
            Alert alert = new Alert(Alert.AlertType.ERROR,
                    "Failed to start application: " + e.getMessage());
            alert.showAndWait();
            Platform.exit();
        }
    }
}
