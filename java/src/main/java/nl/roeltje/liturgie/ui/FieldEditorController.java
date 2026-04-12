package nl.roeltje.liturgie.ui;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.Stage;
import nl.roeltje.liturgie.AppContext;
import nl.roeltje.liturgie.models.LiturgySection;
import nl.roeltje.liturgie.models.LiturgySlide;

import java.util.*;

/**
 * Bulk field editor – shows all common fields across all slides.
 */
public class FieldEditorController implements DialogController {

    @FXML private TableView<FieldRow> table;
    @FXML private TableColumn<FieldRow, String> colField, colValue, colCount;
    @FXML private Button btnApply, btnCancel;

    private final AppContext ctx = AppContext.get();

    public record FieldRow(String name, javafx.beans.property.StringProperty valueProperty, int count) {}

    @FXML
    public void initialize() {
        colField.setCellValueFactory(d -> new javafx.beans.property.SimpleStringProperty(d.getValue().name()));
        colValue.setCellValueFactory(d -> d.getValue().valueProperty());
        colValue.setCellFactory(TextFieldTableCell.forTableColumn());
        colValue.setEditable(true);
        colCount.setCellValueFactory(d -> new javafx.beans.property.SimpleStringProperty(String.valueOf(d.getValue().count())));
        table.setEditable(true);
        btnApply.setText(ctx.tr("dialog.fields.apply"));
        btnCancel.setText(ctx.tr("button.cancel"));

        if (ctx.getCurrentLiturgy() != null) loadFields();
    }

    private void loadFields() {
        Map<String, Integer> fieldCounts = new LinkedHashMap<>();
        Map<String, String> fieldValues = new LinkedHashMap<>();
        for (LiturgySection sec : ctx.getCurrentLiturgy().getSections()) {
            for (LiturgySlide sl : sec.getSlides()) {
                for (Map.Entry<String, String> e : sl.getFields().entrySet()) {
                    fieldCounts.merge(e.getKey(), 1, Integer::sum);
                    fieldValues.putIfAbsent(e.getKey(), e.getValue());
                }
            }
        }
        for (Map.Entry<String, Integer> e : fieldCounts.entrySet()) {
            var prop = new javafx.beans.property.SimpleStringProperty(fieldValues.getOrDefault(e.getKey(), ""));
            table.getItems().add(new FieldRow(e.getKey(), prop, e.getValue()));
        }
    }

    @FXML void onApply() {
        Map<String, String> values = new HashMap<>();
        table.getItems().forEach(row -> values.put(row.name(), row.valueProperty().get()));
        if (ctx.getCurrentLiturgy() != null) {
            for (LiturgySection sec : ctx.getCurrentLiturgy().getSections()) {
                for (LiturgySlide sl : sec.getSlides()) {
                    sl.getFields().putAll(values);
                }
            }
        }
        close();
    }

    @FXML void onCancel() { close(); }

    private void close() { ((Stage) btnCancel.getScene().getWindow()).close(); }
}
