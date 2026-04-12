package nl.roeltje.liturgie.ui;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.Stage;
import nl.roeltje.liturgie.AppContext;
import nl.roeltje.liturgie.models.*;
import nl.roeltje.liturgie.services.SongMatcherService.MatchResult;

import java.util.Arrays;
import java.util.List;

/**
 * Quick-fill songs dialog – paste song names, match to library, add to liturgy.
 */
public class QuickLiturgyController implements DialogController {

    @FXML private TextArea taInput;
    @FXML private TableView<MatchRow> tableResults;
    @FXML private TableColumn<MatchRow, String> colTyped, colMatched;
    @FXML private Button btnMatch, btnAdd, btnCancel;

    private final AppContext ctx = AppContext.get();
    private MainController main;
    private List<MatchResult> lastMatches;

    public record MatchRow(String typed, String matched, MatchResult result) {}

    @Override public void setMainController(MainController main) { this.main = main; }

    @FXML
    public void initialize() {
        colTyped.setCellValueFactory(d -> new javafx.beans.property.SimpleStringProperty(d.getValue().typed()));
        colMatched.setCellValueFactory(d -> new javafx.beans.property.SimpleStringProperty(d.getValue().matched()));
        btnMatch.setText(ctx.tr("dialog.quick_fill.match_btn"));
        btnAdd.setText(ctx.tr("dialog.quick_fill.add_btn"));
        btnCancel.setText(ctx.tr("button.cancel"));
    }

    @FXML void onMatch() {
        tableResults.getItems().clear();
        List<String> lines = Arrays.stream(taInput.getText().split("\n"))
                .map(String::trim).filter(l -> !l.isBlank()).toList();
        if (lines.isEmpty()) return;
        var songs = ctx.getFolderScanner().scanSongs(ctx.getSettings().getSongsPath());
        lastMatches = ctx.getSongMatcher().matchAll(lines, songs);
        for (MatchResult r : lastMatches) {
            String matched = r.matched() != null ? r.matched().getDisplayTitle()
                    : ctx.tr("dialog.quick_fill.stub_label");
            tableResults.getItems().add(new MatchRow(r.input(), matched, r));
        }
    }

    @FXML void onAdd() {
        if (lastMatches == null || ctx.getCurrentLiturgy() == null) return;
        for (MatchResult match : lastMatches) {
            LiturgySection sec = new LiturgySection();
            sec.setSection_type(SectionType.SONG);
            if (match.matched() != null) {
                sec.setName(match.matched().getDisplayTitle());
                if (match.matched().getPdfPath() != null) sec.setPdf_path(match.matched().getPdfPath().toString());
                sec.setYoutube_links(new java.util.ArrayList<>(match.matched().getYoutubeLinks()));
                LiturgySlide sl = new LiturgySlide();
                sl.setTitle(sec.getName());
                sl.setSource_path(match.matched().getPptxPath() != null ? match.matched().getPptxPath().toString() : null);
                sec.getSlides().add(sl);
            } else {
                sec.setName(match.input());
                LiturgySlide sl = new LiturgySlide();
                sl.setTitle(match.input());
                sl.setIs_stub(true);
                sec.getSlides().add(sl);
            }
            ctx.getCurrentLiturgy().addSection(sec);
        }
        if (main != null) main.markDirty();
        close();
    }

    @FXML void onCancel() { close(); }
    private void close() { ((Stage) btnCancel.getScene().getWindow()).close(); }
}
