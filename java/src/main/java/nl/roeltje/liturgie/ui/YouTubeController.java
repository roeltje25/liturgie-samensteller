package nl.roeltje.liturgie.ui;

import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.Stage;
import nl.roeltje.liturgie.AppContext;
import nl.roeltje.liturgie.models.LiturgySection;
import nl.roeltje.liturgie.services.YouTubeService.YouTubeResult;

import java.util.List;

/**
 * YouTube search and link assignment dialog.
 */
public class YouTubeController implements DialogController {

    @FXML private TextField tfSearch;
    @FXML private ListView<YouTubeResult> lstResults;
    @FXML private Label lblStatus;
    @FXML private ProgressIndicator spinner;
    @FXML private Button btnSearch, btnSave, btnCancel;

    private final AppContext ctx = AppContext.get();
    private String sectionId;
    private MainController main;

    @Override public void setArg(String arg) { this.sectionId = arg; }
    @Override public void setMainController(MainController main) { this.main = main; }

    @FXML
    public void initialize() {
        spinner.setVisible(false);
        lstResults.setCellFactory(lv -> new ListCell<>() {
            @Override protected void updateItem(YouTubeResult r, boolean empty) {
                super.updateItem(r, empty);
                setText(empty || r == null ? null : r.title() + "  [" + r.duration() + "]  " + r.channel());
            }
        });
        btnSearch.setText(ctx.tr("button.search"));
        btnSave.setText(ctx.tr("dialog.youtube.save_link"));
        btnCancel.setText(ctx.tr("button.cancel"));

        // Pre-fill search from section name
        if (sectionId != null && ctx.getCurrentLiturgy() != null) {
            ctx.getCurrentLiturgy().findSectionById(sectionId).ifPresent(s -> tfSearch.setText(s.getName()));
        }
    }

    @FXML void onSearch() {
        String query = tfSearch.getText().trim();
        if (query.isBlank()) return;
        spinner.setVisible(true);
        lstResults.getItems().clear();
        lblStatus.setText(ctx.tr("dialog.youtube.searching"));

        Task<List<YouTubeResult>> task = new Task<>() {
            @Override protected List<YouTubeResult> call() {
                return ctx.getYouTubeService().search(query, 10);
            }
        };
        task.setOnSucceeded(e -> {
            List<YouTubeResult> results = task.getValue();
            lstResults.getItems().setAll(results);
            lblStatus.setText(results.isEmpty() ? ctx.tr("dialog.youtube.no_results") : "");
            spinner.setVisible(false);
        });
        task.setOnFailed(e -> { lblStatus.setText("Error"); spinner.setVisible(false); });
        new Thread(task, "yt-search").start();
    }

    @FXML void onSave() {
        YouTubeResult sel = lstResults.getSelectionModel().getSelectedItem();
        if (sel == null || sectionId == null) return;
        LiturgySection sec = ctx.getCurrentLiturgy() != null
                ? ctx.getCurrentLiturgy().findSectionById(sectionId).orElse(null) : null;
        if (sec != null) {
            if (!sec.getYoutube_links().contains(sel.url())) sec.getYoutube_links().add(sel.url());
            if (main != null) main.markDirty();
        }
        close();
    }

    @FXML void onCancel() { close(); }
    private void close() { ((Stage) btnCancel.getScene().getWindow()).close(); }
}
