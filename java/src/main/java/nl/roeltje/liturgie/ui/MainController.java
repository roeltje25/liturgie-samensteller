package nl.roeltje.liturgie.ui;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.cell.TextFieldTreeCell;
import javafx.scene.layout.VBox;
import javafx.stage.*;
import nl.roeltje.liturgie.AppContext;
import nl.roeltje.liturgie.models.*;
import nl.roeltje.liturgie.i18n.TranslationManager;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.*;
import java.time.LocalDate;
import java.util.List;
import java.util.Optional;

/**
 * Controller for the main application window.
 *
 * Responsibilities:
 * - Menu actions (File / Edit / Tools / Help)
 * - Left-panel add-item buttons
 * - Liturgy tree display and selection
 * - Status bar updates
 */
public class MainController {

    private static final Logger log = LoggerFactory.getLogger(MainController.class);

    // ── FXML injections ────────────────────────────────────────────────────────
    @FXML private MenuBar menuBar;
    @FXML private Menu menuFile, menuEdit, menuTools, menuHelp;
    @FXML private MenuItem menuNew, menuOpen, menuSave, menuSaveAs, menuExport;
    @FXML private MenuItem menuOpenTheme, menuSaveTheme, menuExit;
    @FXML private MenuItem menuDelete, menuMoveUp, menuMoveDown;
    @FXML private MenuItem menuCheckLinks, menuEditFields, menuAddBible, menuImportPptx, menuSettings;
    @FXML private MenuItem menuShortcuts, menuWorkflow, menuAbout;

    @FXML private Button btnAddSong, btnQuickFill, btnCreateSong;
    @FXML private Button btnAddGeneric, btnAddOffering, btnAddTheme;
    @FXML private Button btnAddSection, btnAddPptx, btnAddBible;
    @FXML private Button btnEditFields, btnShare;
    @FXML private Button btnDelete, btnEdit;

    @FXML private DatePicker datePicker;
    @FXML private TextField tfLeider;
    @FXML private Label lblWarning, lblStatus, lblAddItems, lblLiturgy;
    @FXML private TreeView<LiturgyNode> liturgyTree;

    private Stage stage;
    private final AppContext ctx = AppContext.get();
    private final TranslationManager tr = TranslationManager.getInstance();
    private boolean unsavedChanges = false;

    // ── Lifecycle ──────────────────────────────────────────────────────────────

    @FXML
    public void initialize() {
        setupTree();
        setupServiceInfoBindings();
        applyTranslations();
        tr.addLanguageChangeListener(this::applyTranslations);
        setStatus(tr.tr("status.ready"));
    }

    public void initStage(Stage stage) {
        this.stage = stage;
        stage.setOnCloseRequest(e -> {
            if (!confirmUnsaved()) e.consume();
            else {
                saveWindowState();
                Platform.exit();
            }
        });
        addLanguageMenu();
    }

    // ── First-run ──────────────────────────────────────────────────────────────

    public void showFirstRunDialog() {
        DirectoryChooser dc = new DirectoryChooser();
        dc.setTitle(tr.tr("dialog.firstrun.title"));
        File dir = dc.showDialog(stage);
        if (dir != null) {
            ctx.getSettings().setBase_folder(dir.getAbsolutePath());
            ctx.getSettings().save();
            ctx.getFolderScanner().refresh();
        }
    }

    // ── Tree setup ─────────────────────────────────────────────────────────────

    private void setupTree() {
        liturgyTree.setRoot(new TreeItem<>(new LiturgyNode("root")));
        liturgyTree.setShowRoot(false);
        liturgyTree.setCellFactory(tv -> new LiturgyTreeCell());
        liturgyTree.setOnMouseClicked(e -> {
            if (e.getClickCount() == 2) onEdit();
        });

        // Context menu
        ContextMenu cm = new ContextMenu();
        MenuItem cmEdit = new MenuItem(tr.tr("button.edit"));
        MenuItem cmDuplicate = new MenuItem(tr.tr("context.duplicate"));
        MenuItem cmDelete = new MenuItem(tr.tr("button.delete"));
        cmEdit.setOnAction(e -> onEdit());
        cmDuplicate.setOnAction(e -> onDuplicate());
        cmDelete.setOnAction(e -> onDelete());
        cm.getItems().addAll(cmEdit, cmDuplicate, new SeparatorMenuItem(), cmDelete);
        liturgyTree.setContextMenu(cm);
    }

    private void setupServiceInfoBindings() {
        datePicker.valueProperty().addListener((obs, old, val) -> {
            Liturgy liturgy = ctx.getCurrentLiturgy();
            if (liturgy != null && val != null) {
                liturgy.setService_date(val.toString());
                markDirty();
            }
        });
        tfLeider.textProperty().addListener((obs, old, val) -> {
            Liturgy liturgy = ctx.getCurrentLiturgy();
            if (liturgy != null) {
                liturgy.setDienstleider(val);
                markDirty();
            }
        });
    }

    private void refreshTree() {
        Liturgy liturgy = ctx.getCurrentLiturgy();
        TreeItem<LiturgyNode> root = new TreeItem<>(new LiturgyNode("root"));
        if (liturgy != null) {
            for (LiturgySection section : liturgy.getSections()) {
                TreeItem<LiturgyNode> sectionItem = new TreeItem<>(
                        new LiturgyNode(section.getName(), section.getId(), true));
                sectionItem.setExpanded(true);
                for (LiturgySlide slide : section.getSlides()) {
                    TreeItem<LiturgyNode> slideItem = new TreeItem<>(
                            new LiturgyNode(slide.getTitle(), slide.getId(), false));
                    sectionItem.getChildren().add(slideItem);
                }
                root.getChildren().add(sectionItem);
            }
            // Service info
            if (liturgy.getService_date() != null) {
                try { datePicker.setValue(LocalDate.parse(liturgy.getService_date())); }
                catch (Exception ignored) {}
            }
            tfLeider.setText(liturgy.getDienstleider() != null ? liturgy.getDienstleider() : "");
        }
        liturgyTree.setRoot(root);
    }

    // ── Translation ────────────────────────────────────────────────────────────

    private void applyTranslations() {
        menuFile.setText(tr.tr("menu.file"));
        menuNew.setText(tr.tr("menu.file.new"));
        menuOpen.setText(tr.tr("menu.file.open"));
        menuSave.setText(tr.tr("menu.file.save"));
        menuSaveAs.setText(tr.tr("menu.file.save_as"));
        menuExport.setText(tr.tr("menu.file.export"));
        menuOpenTheme.setText(tr.tr("menu.file.open_theme"));
        menuSaveTheme.setText(tr.tr("menu.file.save_as_theme"));
        menuExit.setText(tr.tr("menu.file.exit"));
        menuEdit.setText(tr.tr("menu.edit"));
        menuDelete.setText(tr.tr("menu.edit.delete"));
        menuMoveUp.setText(tr.tr("menu.edit.move_up"));
        menuMoveDown.setText(tr.tr("menu.edit.move_down"));
        menuTools.setText(tr.tr("menu.tools"));
        menuCheckLinks.setText(tr.tr("menu.tools.check_links"));
        menuEditFields.setText(tr.tr("menu.tools.edit_fields"));
        menuAddBible.setText(tr.tr("menu.tools.add_bible"));
        menuImportPptx.setText(tr.tr("menu.tools.import_pptx"));
        menuSettings.setText(tr.tr("menu.tools.settings"));
        menuHelp.setText(tr.tr("menu.help"));
        menuShortcuts.setText(tr.tr("menu.help.shortcuts"));
        menuWorkflow.setText(tr.tr("menu.help.workflow"));
        menuAbout.setText(tr.tr("menu.help.about"));

        btnAddSong.setText(tr.tr("button.add_song"));
        btnQuickFill.setText(tr.tr("button.quick_fill_songs"));
        btnCreateSong.setText(tr.tr("button.create_song"));
        btnAddGeneric.setText(tr.tr("button.add_generic"));
        btnAddOffering.setText(tr.tr("button.add_offering"));
        btnAddTheme.setText(tr.tr("button.add_from_theme"));
        btnAddSection.setText(tr.tr("button.add_section"));
        btnAddPptx.setText(tr.tr("button.add_pptx"));
        btnAddBible.setText(tr.tr("button.add_bible"));
        btnEditFields.setText(tr.tr("button.edit_fields"));
        btnShare.setText(tr.tr("button.share"));
        btnDelete.setText(tr.tr("button.delete"));
        btnEdit.setText(tr.tr("button.edit"));
        lblAddItems.setText(tr.tr("panel.available"));
        lblLiturgy.setText(tr.tr("panel.liturgy"));

        if (stage != null) stage.setTitle(tr.tr("app.title"));
    }

    private void addLanguageMenu() {
        Menu langMenu = new Menu("🌐");
        RadioMenuItem nlItem = new RadioMenuItem(tr.tr("language.nl"));
        RadioMenuItem enItem = new RadioMenuItem(tr.tr("language.en"));
        ToggleGroup tg = new ToggleGroup();
        nlItem.setToggleGroup(tg);
        enItem.setToggleGroup(tg);
        (ctx.getSettings().getLanguage().equals("en") ? enItem : nlItem).setSelected(true);
        nlItem.setOnAction(e -> switchLanguage("nl"));
        enItem.setOnAction(e -> switchLanguage("en"));
        langMenu.getItems().addAll(nlItem, enItem);
        menuBar.getMenus().add(langMenu);
    }

    private void switchLanguage(String lang) {
        ctx.getSettings().setLanguage(lang);
        ctx.getSettings().save();
        tr.setLanguage(lang);
    }

    // ── File menu actions ──────────────────────────────────────────────────────

    @FXML void onNew() {
        if (!confirmUnsaved()) return;
        String name = "Liturgie " + LocalDate.now();
        Liturgy liturgy = new Liturgy(name);
        ctx.setCurrentLiturgy(liturgy);
        ctx.setCurrentLiturgyPath(null);
        unsavedChanges = false;
        refreshTree();
        updateTitle();
        setStatus(tr.tr("status.ready"));
    }

    @FXML void onOpen() {
        if (!confirmUnsaved()) return;
        FileChooser fc = new FileChooser();
        fc.setTitle(tr.tr("menu.file.open"));
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Liturgie bestanden (*.liturgy)", "*.liturgy"));
        configureInitialDir(fc);
        File file = fc.showOpenDialog(stage);
        if (file == null) return;
        try {
            Path basePath = ctx.getSettings().isFirstRun() ? null
                    : ctx.getSettings().getSongsPath().getParent();
            Liturgy.MigrationResult result = Liturgy.loadWithMigration(file.toPath(), basePath);
            ctx.setCurrentLiturgy(result.liturgy());
            ctx.setCurrentLiturgyPath(file.getAbsolutePath());
            unsavedChanges = result.wasMigrated();
            refreshTree();
            updateTitle();
            if (result.wasMigrated()) {
                showInfo(tr.tr("dialog.migration.title"), tr.tr("dialog.migration.text"));
            }
        } catch (IOException e) {
            showError(tr.tr("error.load_failed", "error", e.getMessage()));
        }
    }

    @FXML void onSave() {
        if (ctx.getCurrentLiturgy() == null) return;
        if (ctx.getCurrentLiturgyPath() == null) { onSaveAs(); return; }
        saveTo(ctx.getCurrentLiturgyPath());
    }

    @FXML void onSaveAs() {
        if (ctx.getCurrentLiturgy() == null) return;
        FileChooser fc = new FileChooser();
        fc.setTitle(tr.tr("menu.file.save_as"));
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Liturgie bestanden (*.liturgy)", "*.liturgy"));
        configureInitialDir(fc);
        File file = fc.showSaveDialog(stage);
        if (file == null) return;
        String path = file.getAbsolutePath();
        if (!path.endsWith(".liturgy")) path += ".liturgy";
        saveTo(path);
    }

    private void saveTo(String path) {
        try {
            Path basePath = ctx.getSettings().isFirstRun() ? null
                    : ctx.getSettings().getSongsPath().getParent();
            ctx.getCurrentLiturgy().save(Path.of(path), basePath);
            ctx.setCurrentLiturgyPath(path);
            unsavedChanges = false;
            updateTitle();
            setStatus(tr.tr("status.saving") + " → " + Path.of(path).getFileName());
        } catch (IOException e) {
            showError(tr.tr("error.save_failed", "error", e.getMessage()));
        }
    }

    @FXML void onExport() {
        if (ctx.getCurrentLiturgy() == null) return;
        openDialog("export.fxml", tr.tr("dialog.export.title"));
    }

    @FXML void onOpenTheme() {
        FileChooser fc = new FileChooser();
        fc.setTitle(tr.tr("dialog.theme.browse_title"));
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Liturgie bestanden (*.liturgy)", "*.liturgy"));
        File file = fc.showOpenDialog(stage);
        if (file == null) return;
        openDialogWithFile("theme-picker.fxml", tr.tr("dialog.theme.title"), file.getAbsolutePath());
    }

    @FXML void onSaveAsTheme() {
        if (ctx.getCurrentLiturgy() == null) return;
        FileChooser fc = new FileChooser();
        fc.setTitle(tr.tr("menu.file.save_as_theme"));
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Liturgie bestanden (*.liturgy)", "*.liturgy"));
        Path themesDir = ctx.getSettings().getThemesPath();
        if (Files.isDirectory(themesDir)) fc.setInitialDirectory(themesDir.toFile());
        File file = fc.showSaveDialog(stage);
        if (file == null) return;
        try {
            Path path = file.toPath();
            if (!file.getName().endsWith(".liturgy"))
                path = path.resolveSibling(file.getName() + ".liturgy");
            ctx.getThemeService().saveTheme(ctx.getCurrentLiturgy(), path);
            setStatus(tr.tr("dialog.theme.saved", "path", path.getFileName()));
        } catch (IOException e) {
            showError(tr.tr("error.save_failed", "error", e.getMessage()));
        }
    }

    @FXML void onExit() {
        if (confirmUnsaved()) {
            saveWindowState();
            Platform.exit();
        }
    }

    // ── Edit actions ──────────────────────────────────────────────────────────

    @FXML void onDelete() {
        TreeItem<LiturgyNode> selected = liturgyTree.getSelectionModel().getSelectedItem();
        if (selected == null || ctx.getCurrentLiturgy() == null) return;
        LiturgyNode node = selected.getValue();
        if (!confirmDelete(node.name())) return;

        if (node.isSection()) {
            ctx.getCurrentLiturgy().findSectionById(node.id()).ifPresent(sec -> {
                int idx = ctx.getCurrentLiturgy().getSections().indexOf(sec);
                ctx.getCurrentLiturgy().removeSection(idx);
            });
        } else {
            ctx.getCurrentLiturgy().getSections().forEach(sec ->
                    sec.getSlides().removeIf(sl -> sl.getId().equals(node.id())));
        }
        markDirty();
        refreshTree();
    }

    @FXML void onMoveUp() {
        moveSelected(-1);
    }

    @FXML void onMoveDown() {
        moveSelected(1);
    }

    private void moveSelected(int delta) {
        TreeItem<LiturgyNode> selected = liturgyTree.getSelectionModel().getSelectedItem();
        if (selected == null || ctx.getCurrentLiturgy() == null) return;
        LiturgyNode node = selected.getValue();
        List<LiturgySection> sections = ctx.getCurrentLiturgy().getSections();
        if (node.isSection()) {
            int idx = -1;
            for (int i = 0; i < sections.size(); i++) {
                if (sections.get(i).getId().equals(node.id())) { idx = i; break; }
            }
            if (idx >= 0 && idx + delta >= 0 && idx + delta < sections.size()) {
                ctx.getCurrentLiturgy().moveSection(idx, idx + delta);
                markDirty();
                refreshTree();
            }
        }
    }

    private void onDuplicate() {
        TreeItem<LiturgyNode> selected = liturgyTree.getSelectionModel().getSelectedItem();
        if (selected == null || ctx.getCurrentLiturgy() == null) return;
        LiturgyNode node = selected.getValue();
        if (node.isSection()) {
            ctx.getCurrentLiturgy().findSectionById(node.id()).ifPresent(sec -> {
                int idx = ctx.getCurrentLiturgy().getSections().indexOf(sec);
                ctx.getCurrentLiturgy().insertSection(idx + 1, sec.copy());
                markDirty();
                refreshTree();
            });
        }
    }

    // ── Add-item actions ───────────────────────────────────────────────────────

    @FXML void onAddSong()       { openAddDialog("song-picker.fxml",    tr.tr("dialog.song.title")); }
    @FXML void onQuickFill()     { openDialog("quick-liturgy.fxml",     tr.tr("dialog.quick_fill.title")); }
    @FXML void onCreateSong()    { openDialog("new-song.fxml",          tr.tr("dialog.newsong.title")); }
    @FXML void onAddGeneric()    { openAddDialog("generic-picker.fxml", tr.tr("dialog.generic.title")); }
    @FXML void onAddOffering()   { openAddDialog("offering-picker.fxml",tr.tr("dialog.offering.title")); }
    @FXML void onAddFromTheme()  { openDialog("theme-picker.fxml",      tr.tr("dialog.theme.title")); }

    @FXML void onAddSection() {
        TextInputDialog dlg = new TextInputDialog();
        dlg.setTitle(tr.tr("dialog.section.title"));
        dlg.setHeaderText(tr.tr("dialog.section.enter_name"));
        dlg.showAndWait().ifPresent(name -> {
            if (!name.isBlank()) {
                LiturgySection sec = new LiturgySection();
                sec.setName(name.trim());
                ensureLiturgy();
                ctx.getCurrentLiturgy().addSection(sec);
                markDirty();
                refreshTree();
            }
        });
    }

    @FXML void onAddPptx() {
        FileChooser fc = new FileChooser();
        fc.setTitle(tr.tr("dialog.pptx.browse_title"));
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("PowerPoint (*.pptx)", "*.pptx"));
        File file = fc.showOpenDialog(stage);
        if (file == null) return;
        TextInputDialog dlg = new TextInputDialog(file.getName().replace(".pptx", ""));
        dlg.setTitle(tr.tr("dialog.pptx.title"));
        dlg.setHeaderText(tr.tr("dialog.pptx.enter_name"));
        dlg.showAndWait().ifPresent(name -> {
            if (!name.isBlank()) {
                int count = ctx.getPptxService().getSlideCount(file.toPath());
                if (count == 0) { showWarning(tr.tr("dialog.pptx.no_slides")); return; }
                LiturgySection sec = new LiturgySection();
                sec.setName(name.trim());
                for (int i = 0; i < count; i++) {
                    LiturgySlide sl = new LiturgySlide();
                    sl.setSource_path(file.getAbsolutePath());
                    sl.setSlide_index(i);
                    sl.setTitle(name + " - Dia " + (i + 1));
                    sec.getSlides().add(sl);
                }
                ensureLiturgy();
                ctx.getCurrentLiturgy().addSection(sec);
                markDirty();
                refreshTree();
            }
        });
    }

    @FXML void onAddBible()    { openDialog("bible-picker.fxml",     tr.tr("dialog.bible.title")); }
    @FXML void onEditFields()  { openDialog("field-editor.fxml",     tr.tr("dialog.fields.bulk_title")); }
    @FXML void onShare()       { copyLinksToClipboard(); }

    // ── Tools actions ──────────────────────────────────────────────────────────

    @FXML void onCheckLinks()  {
        if (ctx.getCurrentLiturgy() == null) return;
        setStatus(tr.tr("status.checking_links"));
        // Background validation
        javafx.concurrent.Task<java.util.Map<String,Boolean>> task = new javafx.concurrent.Task<>() {
            @Override protected java.util.Map<String,Boolean> call() {
                java.util.List<String> urls = new java.util.ArrayList<>();
                ctx.getCurrentLiturgy().getSections().forEach(s -> urls.addAll(s.getYoutube_links()));
                return ctx.getYouTubeService().validateBatch(urls);
            }
        };
        task.setOnSucceeded(e -> {
            long invalid = task.getValue().values().stream().filter(v -> !v).count();
            setStatus(invalid == 0 ? tr.tr("status.links_valid")
                    : tr.tr("status.links_invalid", "count", invalid));
        });
        new Thread(task, "check-links").start();
    }

    @FXML void onImportPptx()  { openDialog("import-pptx.fxml",     tr.tr("dialog.import_pptx.title")); }
    @FXML void onSettings()    { openDialog("settings.fxml",         tr.tr("dialog.settings.title")); }

    // ── Help actions ───────────────────────────────────────────────────────────

    @FXML void onShortcuts() {
        Alert a = new Alert(Alert.AlertType.INFORMATION);
        a.setTitle(tr.tr("dialog.shortcuts.title"));
        a.setHeaderText(null);
        a.setContentText(tr.tr("dialog.shortcuts.content")
                .replaceAll("<[^>]+>", "").replace("&gt;", ">"));
        a.showAndWait();
    }

    @FXML void onWorkflow() {
        Alert a = new Alert(Alert.AlertType.INFORMATION);
        a.setTitle(tr.tr("dialog.workflow.title"));
        a.setHeaderText(null);
        a.setContentText(tr.tr("dialog.workflow.content")
                .replaceAll("<[^>]+>", "").replace("&gt;", ">"));
        a.showAndWait();
    }

    @FXML void onAbout() { openDialog("about.fxml", tr.tr("dialog.about.title")); }

    @FXML void onEdit() {
        TreeItem<LiturgyNode> selected = liturgyTree.getSelectionModel().getSelectedItem();
        if (selected == null) return;
        LiturgyNode node = selected.getValue();
        if (node.isSection()) {
            openDialogWithFile("section-editor.fxml", tr.tr("dialog.editor.title"), node.id());
        } else {
            openDialogWithFile("field-editor.fxml", tr.tr("dialog.fields.slide_title", "slide", node.name()), node.id());
        }
    }

    // ── Helpers ────────────────────────────────────────────────────────────────

    private void openAddDialog(String fxmlName, String title) {
        ensureLiturgy();
        openDialog(fxmlName, title);
    }

    private void openDialog(String fxmlName, String title) {
        openDialogWithFile(fxmlName, title, null);
    }

    private void openDialogWithFile(String fxmlName, String title, String arg) {
        try {
            var url = getClass().getResource("/nl/roeltje/liturgie/fxml/" + fxmlName);
            if (url == null) { showError("FXML not found: " + fxmlName); return; }
            FXMLLoader loader = new FXMLLoader(url);
            var root = loader.load();
            if (loader.getController() instanceof DialogController dc) {
                dc.setArg(arg);
                dc.setMainController(this);
            }
            Stage dlgStage = new Stage();
            dlgStage.setTitle(title);
            dlgStage.initModality(Modality.WINDOW_MODAL);
            dlgStage.initOwner(stage);
            dlgStage.setScene(new Scene(root));
            dlgStage.showAndWait();
            refreshTree();
            markDirty();
        } catch (IOException e) {
            log.error("Cannot open dialog {}", fxmlName, e);
            showError(e.getMessage());
        }
    }

    private void ensureLiturgy() {
        if (ctx.getCurrentLiturgy() == null) {
            onNew();
        }
    }

    private void copyLinksToClipboard() {
        if (ctx.getCurrentLiturgy() == null) return;
        String text = ctx.getExportService().generateLinksText(ctx.getCurrentLiturgy());
        javafx.scene.input.ClipboardContent cc = new javafx.scene.input.ClipboardContent();
        cc.putString(text);
        javafx.scene.input.Clipboard.getSystemClipboard().setContent(cc);
        setStatus(tr.tr("status.share_clipboard"));
    }

    private boolean confirmUnsaved() {
        if (!unsavedChanges || ctx.getCurrentLiturgy() == null) return true;
        Alert a = new Alert(Alert.AlertType.CONFIRMATION,
                tr.tr("dialog.confirm.unsaved_text"),
                ButtonType.YES, ButtonType.NO, ButtonType.CANCEL);
        a.setTitle(tr.tr("dialog.confirm.unsaved"));
        Optional<ButtonType> result = a.showAndWait();
        if (result.isEmpty() || result.get() == ButtonType.CANCEL) return false;
        if (result.get() == ButtonType.YES) onSave();
        return true;
    }

    private boolean confirmDelete(String name) {
        Alert a = new Alert(Alert.AlertType.CONFIRMATION,
                tr.tr("dialog.confirm.delete_text", "title", name));
        a.setTitle(tr.tr("dialog.confirm.delete"));
        return a.showAndWait().filter(b -> b == ButtonType.OK).isPresent();
    }

    public void markDirty() {
        unsavedChanges = true;
        updateTitle();
    }

    private void updateTitle() {
        if (stage == null) return;
        Liturgy l = ctx.getCurrentLiturgy();
        String name = l != null ? l.getName() : tr.tr("app.title");
        stage.setTitle((unsavedChanges ? "* " : "") + name + " – " + tr.tr("app.title"));
    }

    public void setStatus(String message) {
        if (lblStatus != null) lblStatus.setText(message);
    }

    private void configureInitialDir(FileChooser fc) {
        Path output = ctx.getSettings().getOutputPath();
        if (Files.isDirectory(output)) fc.setInitialDirectory(output.toFile());
    }

    private void saveWindowState() {
        if (stage != null) {
            ctx.getSettings().setWindowWidth((int) stage.getWidth());
            ctx.getSettings().setWindowHeight((int) stage.getHeight());
            ctx.getSettings().save();
        }
    }

    private void showError(String msg) {
        Alert a = new Alert(Alert.AlertType.ERROR, msg);
        a.showAndWait();
    }
    private void showWarning(String msg) {
        Alert a = new Alert(Alert.AlertType.WARNING, msg);
        a.showAndWait();
    }
    private void showInfo(String title, String msg) {
        Alert a = new Alert(Alert.AlertType.INFORMATION, msg);
        a.setTitle(title);
        a.showAndWait();
    }

    // ── Inner types ────────────────────────────────────────────────────────────

    /** Data object stored in each TreeItem. */
    public record LiturgyNode(String name, String id, boolean isSection) {
        LiturgyNode(String name) { this(name, null, false); }
        @Override public String toString() { return name; }
    }

    /** Custom tree cell. */
    private static class LiturgyTreeCell extends TreeCell<LiturgyNode> {
        @Override protected void updateItem(LiturgyNode item, boolean empty) {
            super.updateItem(item, empty);
            if (empty || item == null) {
                setText(null);
                setStyle("");
            } else {
                setText(item.name());
                setStyle(item.isSection() ? "-fx-font-weight: bold;" : "");
            }
        }
    }
}
