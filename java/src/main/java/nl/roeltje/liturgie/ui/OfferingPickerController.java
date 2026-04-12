package nl.roeltje.liturgie.ui;

import javafx.application.Platform;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.Stage;
import nl.roeltje.liturgie.AppContext;
import nl.roeltje.liturgie.models.*;
import nl.roeltje.liturgie.services.FolderScannerService.OfferingSlide;

import java.util.List;

/**
 * Offering (collecte) slide picker.
 */
public class OfferingPickerController implements DialogController {

    @FXML private ListView<OfferingSlide> listView;
    @FXML private Label lblPreview;
    @FXML private Button btnSelect, btnCancel;
    @FXML private ProgressIndicator spinner;

    private final AppContext ctx = AppContext.get();
    private MainController main;
    private List<OfferingSlide> slides;

    @Override public void setMainController(MainController main) { this.main = main; }

    @FXML
    public void initialize() {
        spinner.setVisible(true);
        listView.setCellFactory(lv -> new ListCell<>() {
            @Override protected void updateItem(OfferingSlide item, boolean empty) {
                super.updateItem(item, empty);
                setText(empty || item == null ? null : item.title());
            }
        });
        listView.getSelectionModel().selectedItemProperty().addListener((obs, old, sel) -> {
            lblPreview.setText(sel != null ? sel.title() : "");
        });
        listView.setOnMouseClicked(e -> { if (e.getClickCount() == 2) onSelect(); });
        btnSelect.setText(ctx.tr("button.select"));
        btnCancel.setText(ctx.tr("button.cancel"));

        Task<List<OfferingSlide>> task = new Task<>() {
            @Override protected List<OfferingSlide> call() {
                return ctx.getFolderScanner().getOfferingSlides(ctx.getSettings().getCollectePath());
            }
        };
        task.setOnSucceeded(e -> {
            slides = task.getValue();
            listView.getItems().addAll(slides);
            spinner.setVisible(false);
        });
        new Thread(task, "load-offerings").start();
    }

    @FXML void onSelect() {
        OfferingSlide sel = listView.getSelectionModel().getSelectedItem();
        if (sel == null || ctx.getCurrentLiturgy() == null) return;
        LiturgySection sec = new LiturgySection();
        sec.setName(sel.title());
        LiturgySlide sl = new LiturgySlide();
        sl.setSource_path(sel.sourcePath().toString());
        sl.setSlide_index(sel.slideIndex());
        sl.setTitle(sel.title());
        sec.getSlides().add(sl);
        ctx.getCurrentLiturgy().addSection(sec);
        if (main != null) main.markDirty();
        close();
    }

    @FXML void onCancel() { close(); }
    private void close() { ((Stage) btnCancel.getScene().getWindow()).close(); }
}
