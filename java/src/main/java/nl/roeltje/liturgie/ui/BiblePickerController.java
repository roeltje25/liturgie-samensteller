package nl.roeltje.liturgie.ui;

import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.Stage;
import nl.roeltje.liturgie.AppContext;
import nl.roeltje.liturgie.models.*;
import nl.roeltje.liturgie.services.BibleService.*;

import java.nio.file.Path;
import java.util.List;

/**
 * Bible text slides picker.
 */
public class BiblePickerController implements DialogController {

    @FXML private Label lblApiKeyWarning;
    @FXML private TextField tfReference;
    @FXML private ListView<BibleTranslation> lstAvailable, lstSelected;
    @FXML private ComboBox<String> cbLanguage;
    @FXML private TextArea taPreview;
    @FXML private Spinner<Integer> spFontSize, spCharsPerSlide;
    @FXML private TextField tfFontName;
    @FXML private ProgressIndicator spinner;
    @FXML private Label lblStatus;
    @FXML private Button btnGenerate, btnCancel;

    private final AppContext ctx = AppContext.get();
    private MainController main;

    @Override public void setMainController(MainController main) { this.main = main; }

    @FXML
    public void initialize() {
        boolean hasKey = !ctx.getSettings().getYouversion_api_key().isBlank();
        lblApiKeyWarning.setVisible(!hasKey);
        lblApiKeyWarning.setManaged(!hasKey);
        lblApiKeyWarning.setText(ctx.tr("dialog.bible.no_api_key_warning"));

        tfFontName.setText(ctx.getSettings().getBible_font_name());
        spFontSize.setValueFactory(new SpinnerValueFactory.IntegerSpinnerValueFactory(
                6, 72, ctx.getSettings().getBible_font_size()));
        spCharsPerSlide.setValueFactory(new SpinnerValueFactory.IntegerSpinnerValueFactory(
                100, 2000, ctx.getSettings().getBible_chars_per_slide()));

        spinner.setVisible(false);
        btnGenerate.setText(ctx.tr("dialog.bible.generate"));
        btnCancel.setText(ctx.tr("button.cancel"));

        if (hasKey) loadTranslations("nl");
    }

    private void loadTranslations(String lang) {
        spinner.setVisible(true);
        Task<List<BibleTranslation>> task = new Task<>() {
            @Override protected List<BibleTranslation> call() {
                return ctx.getBibleService().getTranslations(lang);
            }
        };
        task.setOnSucceeded(e -> {
            lstAvailable.getItems().setAll(task.getValue());
            spinner.setVisible(false);
        });
        task.setOnFailed(e -> spinner.setVisible(false));
        lstAvailable.setCellFactory(lv -> new ListCell<>() {
            @Override protected void updateItem(BibleTranslation t, boolean empty) {
                super.updateItem(t, empty);
                setText(empty || t == null ? null : t.abbreviation() + " – " + t.name());
            }
        });
        lstSelected.setCellFactory(lstAvailable.getCellFactory());
        new Thread(task, "load-translations").start();
    }

    @FXML void onAddTranslation() {
        BibleTranslation sel = lstAvailable.getSelectionModel().getSelectedItem();
        if (sel != null && lstSelected.getItems().size() < 6 && !lstSelected.getItems().contains(sel)) {
            lstSelected.getItems().add(sel);
        }
    }

    @FXML void onRemoveTranslation() {
        BibleTranslation sel = lstSelected.getSelectionModel().getSelectedItem();
        if (sel != null) lstSelected.getItems().remove(sel);
    }

    @FXML void onFetchPreview() {
        String ref = tfReference.getText().trim();
        if (ref.isBlank() || lstSelected.getItems().isEmpty()) return;
        spinner.setVisible(true);
        BibleTranslation first = lstSelected.getItems().get(0);
        Task<List<BibleVerse>> task = new Task<>() {
            @Override protected List<BibleVerse> call() {
                return ctx.getBibleService().fetchVerses(ref, first.id());
            }
        };
        task.setOnSucceeded(e -> {
            StringBuilder sb = new StringBuilder();
            task.getValue().forEach(v -> sb.append(v.verse()).append(" ").append(v.text()).append("\n"));
            taPreview.setText(sb.toString());
            spinner.setVisible(false);
        });
        task.setOnFailed(e -> { lblStatus.setText(e.getSource().getException().getMessage()); spinner.setVisible(false); });
        new Thread(task, "bible-preview").start();
    }

    @FXML void onGenerate() {
        String ref = tfReference.getText().trim();
        if (ref.isBlank()) { lblStatus.setText(ctx.tr("dialog.bible.error.no_reference")); return; }
        if (lstSelected.getItems().isEmpty()) { lblStatus.setText(ctx.tr("dialog.bible.error.no_translations")); return; }

        spinner.setVisible(true);
        btnGenerate.setDisable(true);
        BibleTranslation first = lstSelected.getItems().get(0);

        Task<Path> task = new Task<>() {
            @Override protected Path call() throws Exception {
                var verses = ctx.getBibleService().fetchVerses(ref, first.id());
                var abbrevs = lstSelected.getItems().stream().map(BibleTranslation::abbreviation).toList();
                return ctx.getBibleSlideService().generateSlides(verses, abbrevs, ref);
            }
        };
        task.setOnSucceeded(e -> {
            Path pptx = task.getValue();
            if (pptx != null && ctx.getCurrentLiturgy() != null) {
                LiturgySection sec = new LiturgySection();
                sec.setName(ref + " (" + first.abbreviation() + ")");
                int count = ctx.getPptxService().getSlideCount(pptx);
                for (int i = 0; i < count; i++) {
                    LiturgySlide sl = new LiturgySlide();
                    sl.setSource_path(pptx.toString());
                    sl.setSlide_index(i);
                    sl.setTitle(ref + " - Dia " + (i + 1));
                    sec.getSlides().add(sl);
                }
                ctx.getCurrentLiturgy().addSection(sec);
                if (main != null) main.markDirty();
                lblStatus.setText(ctx.tr("status.bible_added", "name", ref));
            }
            spinner.setVisible(false);
            btnGenerate.setDisable(false);
        });
        task.setOnFailed(e -> {
            lblStatus.setText(ctx.tr("dialog.bible.error.generation", "error", task.getException().getMessage()));
            spinner.setVisible(false);
            btnGenerate.setDisable(false);
        });
        new Thread(task, "bible-generate").start();
    }

    @FXML void onCancel() { ((Stage) btnCancel.getScene().getWindow()).close(); }
}
