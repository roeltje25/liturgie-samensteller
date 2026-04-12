package nl.roeltje.liturgie.ui;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.Stage;
import nl.roeltje.liturgie.AppContext;
import nl.roeltje.liturgie.models.LiturgySection;

import java.util.ArrayList;
import java.util.List;

/**
 * Section editor dialog – fields tab + YouTube links tab.
 */
public class SectionEditorController implements DialogController {

    @FXML private TabPane tabPane;
    @FXML private TextField tfSectionName;
    @FXML private TableView<FieldRow> fieldsTable;
    @FXML private TableColumn<FieldRow, String> colField, colValue;
    @FXML private ListView<String> lstYoutube;
    @FXML private TextField tfNewYoutubeLink;
    @FXML private Button btnAddLink, btnRemoveLink, btnSave, btnCancel;

    private final AppContext ctx = AppContext.get();
    private String sectionId;
    private LiturgySection section;

    public record FieldRow(String name, String value) {}

    @Override public void setArg(String arg) { this.sectionId = arg; }

    @FXML
    public void initialize() {
        colField.setCellValueFactory(d -> new javafx.beans.property.SimpleStringProperty(d.getValue().name()));
        colValue.setCellValueFactory(d -> new javafx.beans.property.SimpleStringProperty(d.getValue().value()));
        btnSave.setText(ctx.tr("button.save"));
        btnCancel.setText(ctx.tr("button.cancel"));
    }

    public void initData() {
        if (sectionId == null || ctx.getCurrentLiturgy() == null) return;
        section = ctx.getCurrentLiturgy().findSectionById(sectionId).orElse(null);
        if (section == null) return;
        tfSectionName.setText(section.getName());
        // Fields from first slide
        if (!section.getSlides().isEmpty()) {
            section.getSlides().get(0).getFields().forEach((k, v) ->
                    fieldsTable.getItems().add(new FieldRow(k, v)));
        }
        // YouTube links
        lstYoutube.getItems().addAll(section.getYoutube_links());
    }

    @FXML void onAddLink() {
        String url = tfNewYoutubeLink.getText().trim();
        if (!url.isBlank()) {
            lstYoutube.getItems().add(url);
            tfNewYoutubeLink.clear();
        }
    }

    @FXML void onRemoveLink() {
        int idx = lstYoutube.getSelectionModel().getSelectedIndex();
        if (idx >= 0) lstYoutube.getItems().remove(idx);
    }

    @FXML void onSave() {
        if (section != null) {
            section.setName(tfSectionName.getText().trim());
            section.setYoutube_links(new ArrayList<>(lstYoutube.getItems()));
            // Write fields back to all slides
            section.getSlides().forEach(sl -> {
                sl.getFields().clear();
                fieldsTable.getItems().forEach(row -> sl.getFields().put(row.name(), row.value()));
            });
        }
        close();
    }

    @FXML void onCancel() { close(); }
    private void close() { ((Stage) btnCancel.getScene().getWindow()).close(); }
}
