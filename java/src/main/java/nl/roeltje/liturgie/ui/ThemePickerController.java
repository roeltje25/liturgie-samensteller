package nl.roeltje.liturgie.ui;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import nl.roeltje.liturgie.AppContext;
import nl.roeltje.liturgie.models.*;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;

/**
 * Theme section picker – load sections from a .liturgy theme file.
 */
public class ThemePickerController implements DialogController {

    @FXML private Label lblThemeFile;
    @FXML private TreeView<Object> themeTree;
    @FXML private Button btnBrowse, btnAdd, btnCancel;
    @FXML private Label lblStatus;

    private final AppContext ctx = AppContext.get();
    private MainController main;
    private Liturgy theme;
    private Path themeFile;

    @Override public void setArg(String arg) {
        if (arg != null) loadTheme(Path.of(arg));
    }
    @Override public void setMainController(MainController main) { this.main = main; }

    @FXML
    public void initialize() {
        themeTree.setShowRoot(false);
        themeTree.setCellFactory(tv -> new TreeCell<>() {
            @Override protected void updateItem(Object item, boolean empty) {
                super.updateItem(item, empty);
                if (empty || item == null) { setText(null); }
                else if (item instanceof LiturgySection s) { setText(s.getName()); setStyle("-fx-font-weight: bold;"); }
                else if (item instanceof LiturgySlide sl) { setText(sl.getTitle()); setStyle(""); }
            }
        });
        themeTree.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        btnBrowse.setText(ctx.tr("button.browse_file"));
        btnAdd.setText(ctx.tr("button.add"));
        btnCancel.setText(ctx.tr("button.cancel"));
    }

    @FXML void onBrowse() {
        FileChooser fc = new FileChooser();
        fc.setTitle(ctx.tr("dialog.theme.browse_title"));
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Liturgie bestanden (*.liturgy)", "*.liturgy"));
        Path themesDir = ctx.getSettings().getThemesPath();
        if (java.nio.file.Files.isDirectory(themesDir)) fc.setInitialDirectory(themesDir.toFile());
        java.io.File file = fc.showOpenDialog(btnBrowse.getScene().getWindow());
        if (file != null) loadTheme(file.toPath());
    }

    private void loadTheme(Path path) {
        try {
            themeFile = path;
            theme = ctx.getThemeService().loadTheme(path);
            lblThemeFile.setText(path.getFileName().toString());
            populateTree();
        } catch (IOException e) {
            lblStatus.setText(ctx.tr("error.load_failed", "error", e.getMessage()));
        }
    }

    private void populateTree() {
        TreeItem<Object> root = new TreeItem<>();
        for (LiturgySection sec : theme.getSections()) {
            TreeItem<Object> secItem = new TreeItem<>(sec);
            secItem.setExpanded(true);
            for (LiturgySlide sl : sec.getSlides()) {
                secItem.getChildren().add(new TreeItem<>(sl));
            }
            root.getChildren().add(secItem);
        }
        themeTree.setRoot(root);
    }

    @FXML void onAdd() {
        if (theme == null || ctx.getCurrentLiturgy() == null) return;
        for (TreeItem<Object> item : themeTree.getSelectionModel().getSelectedItems()) {
            if (item.getValue() instanceof LiturgySection sec) {
                ctx.getCurrentLiturgy().addSection(sec.copy());
            }
        }
        if (main != null) main.markDirty();
        close();
    }

    @FXML void onCancel() { close(); }
    private void close() { ((Stage) btnCancel.getScene().getWindow()).close(); }
}
