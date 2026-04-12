package nl.roeltje.liturgie.ui;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.image.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import nl.roeltje.liturgie.AppContext;
import nl.roeltje.liturgie.models.*;
import nl.roeltje.liturgie.services.FolderScannerService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.image.BufferedImage;
import java.io.File;
import java.nio.file.*;
import java.util.List;
import java.util.Optional;
import java.util.concurrent.CompletableFuture;

/**
 * Song picker dialog.
 * Shows a tree of songs grouped by folder, with live thumbnail preview.
 */
public class SongPickerController implements DialogController {

    private static final Logger log = LoggerFactory.getLogger(SongPickerController.class);

    @FXML private TreeView<Song> songTree;
    @FXML private TextField tfSearch;
    @FXML private ImageView imgPreview;
    @FXML private Label lblInfo;
    @FXML private Button btnSelect, btnCancel, btnBrowse, btnStub;

    private final AppContext ctx = AppContext.get();
    private MainController main;
    private List<Song> songs;

    @Override public void setMainController(MainController main) { this.main = main; }

    @FXML
    public void initialize() {
        songs = ctx.getFolderScanner().scanSongs(ctx.getSettings().getSongsPath());
        populateTree(songs);

        tfSearch.textProperty().addListener((obs, old, val) -> filterTree(val));

        songTree.getSelectionModel().selectedItemProperty().addListener((obs, old, selected) -> {
            if (selected != null && selected.isLeaf() && selected.getValue() != null) {
                loadThumbnail(selected.getValue());
                updateInfo(selected.getValue());
            }
        });

        btnSelect.setText(ctx.tr("button.select"));
        btnCancel.setText(ctx.tr("button.cancel"));
        btnBrowse.setText(ctx.tr("button.browse_file"));
        btnStub.setText(ctx.tr("button.create_stub"));
    }

    private void populateTree(List<Song> list) {
        TreeItem<Song> root = new TreeItem<>();
        root.setExpanded(true);
        for (Song song : list) {
            // Group by subfolder
            String rel = song.getRelativePath();
            String[] parts = rel.split("[/\\\\]");
            TreeItem<Song> parent = root;
            for (int i = 0; i < parts.length - 1; i++) {
                final String part = parts[i];
                Optional<TreeItem<Song>> group = parent.getChildren().stream()
                        .filter(ti -> ti.getValue() == null
                                && part.equals(ti.getGraphic() instanceof Label l ? l.getText() : ""))
                        .findFirst();
                if (group.isPresent()) {
                    parent = group.get();
                } else {
                    TreeItem<Song> folder = new TreeItem<>();
                    folder.setGraphic(new Label(part));
                    folder.setExpanded(true);
                    parent.getChildren().add(folder);
                    parent = folder;
                }
            }
            parent.getChildren().add(new TreeItem<>(song));
        }
        songTree.setRoot(root);
        songTree.setShowRoot(false);
        songTree.setCellFactory(tv -> new TreeCell<>() {
            @Override protected void updateItem(Song item, boolean empty) {
                super.updateItem(item, empty);
                if (empty) { setText(null); } else if (item == null) {
                    setText(getTreeItem().getGraphic() instanceof Label l ? l.getText() : "");
                    setGraphic(null);
                } else {
                    setText(item.getDisplayTitle());
                }
            }
        });
    }

    private void filterTree(String query) {
        if (query == null || query.isBlank()) {
            populateTree(songs);
            return;
        }
        String lq = query.toLowerCase();
        List<Song> filtered = songs.stream()
                .filter(s -> s.getDisplayTitle().toLowerCase().contains(lq))
                .toList();
        populateTree(filtered);
    }

    private void loadThumbnail(Song song) {
        imgPreview.setImage(null);
        if (song.getPptxPath() == null) return;
        CompletableFuture.supplyAsync(() ->
                ctx.getPptxService().getThumbnail(song.getPptxPath())
        ).thenAccept(optImg -> optImg.ifPresent(img ->
                Platform.runLater(() -> imgPreview.setImage(bufferedToFx(img)))
        ));
    }

    private void updateInfo(Song song) {
        StringBuilder sb = new StringBuilder();
        if (!song.hasPptx()) sb.append(ctx.tr("dialog.song.no_pptx")).append("\n");
        if (!song.hasPdf())  sb.append(ctx.tr("dialog.song.no_pdf")).append("\n");
        if (song.hasYoutube()) sb.append(ctx.tr("dialog.song.has_youtube")).append("\n");
        lblInfo.setText(sb.toString().trim());
    }

    @FXML void onSelect() {
        TreeItem<Song> selected = songTree.getSelectionModel().getSelectedItem();
        if (selected == null || !selected.isLeaf() || selected.getValue() == null) return;
        addSongToLiturgy(selected.getValue());
        close();
    }

    @FXML void onBrowse() {
        FileChooser fc = new FileChooser();
        fc.setTitle(ctx.tr("dialog.song.browse_title"));
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("PowerPoint (*.pptx)", "*.pptx"));
        File file = fc.showOpenDialog(btnBrowse.getScene().getWindow());
        if (file == null) return;
        TextInputDialog dlg = new TextInputDialog(file.getName().replace(".pptx", ""));
        dlg.setHeaderText(ctx.tr("dialog.song.external_enter_title"));
        dlg.showAndWait().ifPresent(title -> {
            Song song = new Song(title, file.toPath().getParent(), "");
            song.setTitle(title);
            song.setPptxPath(file.toPath());
            addSongToLiturgy(song);
            close();
        });
    }

    @FXML void onCreateStub() {
        TextInputDialog dlg = new TextInputDialog();
        dlg.setTitle(ctx.tr("dialog.song.stub_dialog_title"));
        dlg.setHeaderText(ctx.tr("dialog.song.stub_enter_title"));
        dlg.showAndWait().ifPresent(title -> {
            if (!title.isBlank()) {
                Song song = new Song(title, null, "");
                song.setTitle(title);
                LiturgySection sec = buildSongSection(song);
                sec.getSlides().get(0).setIs_stub(true);
                ctx.getCurrentLiturgy().addSection(sec);
                main.markDirty();
                close();
            }
        });
    }

    @FXML void onCancel() { close(); }

    private void addSongToLiturgy(Song song) {
        if (main == null) return;
        main.markDirty();
        if (ctx.getCurrentLiturgy() == null) return;
        ctx.getCurrentLiturgy().addSection(buildSongSection(song));
    }

    private LiturgySection buildSongSection(Song song) {
        LiturgySection sec = new LiturgySection();
        sec.setSection_type(SectionType.SONG);
        sec.setName(song.getDisplayTitle());
        if (song.getPdfPath() != null) sec.setPdf_path(song.getPdfPath().toString());
        sec.setYoutube_links(new java.util.ArrayList<>(song.getYoutubeLinks()));

        LiturgySlide slide = new LiturgySlide();
        slide.setTitle(song.getDisplayTitle());
        slide.setSource_path(song.getPptxPath() != null ? song.getPptxPath().toString() : null);
        if (song.getPdfPath() != null) slide.setPdf_path(song.getPdfPath().toString());
        slide.setYoutube_links(new java.util.ArrayList<>(song.getYoutubeLinks()));
        sec.getSlides().add(slide);
        return sec;
    }

    private void close() {
        ((Stage) btnCancel.getScene().getWindow()).close();
    }

    private static Image bufferedToFx(BufferedImage bi) {
        return javafx.embed.swing.SwingFXUtils.toFXImage(bi, null);
    }
}
