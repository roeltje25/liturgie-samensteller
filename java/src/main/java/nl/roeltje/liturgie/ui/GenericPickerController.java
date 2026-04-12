package nl.roeltje.liturgie.ui;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.Stage;
import nl.roeltje.liturgie.AppContext;
import nl.roeltje.liturgie.models.*;

import java.nio.file.Path;
import java.util.List;

/**
 * Generic item (Algemeen) picker dialog.
 */
public class GenericPickerController implements DialogController {

    @FXML private ListView<Path> listView;
    @FXML private TextField tfSearch;
    @FXML private Button btnSelect, btnCancel;

    private final AppContext ctx = AppContext.get();
    private MainController main;
    private List<Path> items;

    @Override public void setMainController(MainController main) { this.main = main; }

    @FXML
    public void initialize() {
        items = ctx.getFolderScanner().scanGeneral(ctx.getSettings().getAlgemeenPath());
        listView.getItems().addAll(items);
        listView.setCellFactory(lv -> new ListCell<>() {
            @Override protected void updateItem(Path item, boolean empty) {
                super.updateItem(item, empty);
                setText(empty || item == null ? null : item.getFileName().toString().replace(".pptx", ""));
            }
        });
        tfSearch.textProperty().addListener((obs, old, val) -> filter(val));
        listView.setOnMouseClicked(e -> { if (e.getClickCount() == 2) onSelect(); });
        btnSelect.setText(ctx.tr("button.select"));
        btnCancel.setText(ctx.tr("button.cancel"));
    }

    private void filter(String query) {
        listView.getItems().setAll(items.stream()
                .filter(p -> query == null || query.isBlank()
                        || p.getFileName().toString().toLowerCase().contains(query.toLowerCase()))
                .toList());
    }

    @FXML void onSelect() {
        Path selected = listView.getSelectionModel().getSelectedItem();
        if (selected == null || ctx.getCurrentLiturgy() == null) return;
        int count = ctx.getPptxService().getSlideCount(selected);
        LiturgySection sec = new LiturgySection();
        sec.setName(selected.getFileName().toString().replace(".pptx", ""));
        for (int i = 0; i < Math.max(1, count); i++) {
            LiturgySlide sl = new LiturgySlide();
            sl.setSource_path(selected.toString());
            sl.setSlide_index(i);
            sl.setTitle(sec.getName() + (count > 1 ? " - Dia " + (i + 1) : ""));
            sec.getSlides().add(sl);
        }
        ctx.getCurrentLiturgy().addSection(sec);
        if (main != null) main.markDirty();
        close();
    }

    @FXML void onCancel() { close(); }
    private void close() { ((Stage) btnCancel.getScene().getWindow()).close(); }
}
