package nl.roeltje.liturgie.ui;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.Stage;
import nl.roeltje.liturgie.AppContext;
import nl.roeltje.liturgie.models.*;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.*;
import java.util.Arrays;
import java.util.List;

/**
 * Create a new song PPTX from lyrics text.
 * Each paragraph (separated by blank line) becomes one slide.
 */
public class NewSongController implements DialogController {

    @FXML private TextField tfTitle, tfSubfolder;
    @FXML private TextArea taLyrics;
    @FXML private Label lblPreview;
    @FXML private Button btnCreate, btnCancel;

    private final AppContext ctx = AppContext.get();
    private MainController main;

    @Override public void setMainController(MainController main) { this.main = main; }

    @FXML
    public void initialize() {
        taLyrics.textProperty().addListener((obs, old, val) -> updatePreview(val));
        btnCreate.setText(ctx.tr("dialog.newsong.create"));
        btnCancel.setText(ctx.tr("button.cancel"));
    }

    private void updatePreview(String text) {
        int count = splitLyrics(text).size();
        lblPreview.setText(ctx.tr("dialog.newsong.preview", "count", count));
    }

    @FXML void onCreate() {
        String title = tfTitle.getText().trim();
        if (title.isBlank()) return;
        List<String> blocks = splitLyrics(taLyrics.getText());
        if (blocks.isEmpty()) return;

        Path songsDir = ctx.getSettings().getSongsPath();
        String subfolder = tfSubfolder.getText().trim();
        Path songDir = subfolder.isBlank() ? songsDir.resolve(title) : songsDir.resolve(subfolder).resolve(title);

        try {
            if (Files.exists(songDir)) {
                Alert a = new Alert(Alert.AlertType.CONFIRMATION,
                        ctx.tr("dialog.newsong.folder_exists", "name", title));
                if (a.showAndWait().filter(b -> b == ButtonType.OK).isEmpty()) return;
            }
            Files.createDirectories(songDir);
            Path pptxPath = songDir.resolve(title + ".pptx");
            generatePptx(pptxPath, title, blocks);

            // Build section
            LiturgySection sec = new LiturgySection();
            sec.setSection_type(SectionType.SONG);
            sec.setName(title);
            for (int i = 0; i < blocks.size(); i++) {
                LiturgySlide sl = new LiturgySlide();
                sl.setSource_path(pptxPath.toString());
                sl.setSlide_index(i);
                sl.setTitle(title + " - Dia " + (i + 1));
                sec.getSlides().add(sl);
            }
            if (ctx.getCurrentLiturgy() != null) {
                ctx.getCurrentLiturgy().addSection(sec);
            }
            ctx.getFolderScanner().refresh();
            if (main != null) main.markDirty();
            close();
        } catch (IOException e) {
            new Alert(Alert.AlertType.ERROR,
                    ctx.tr("dialog.newsong.error", "error", e.getMessage())).showAndWait();
        }
    }

    private void generatePptx(Path output, String title, List<String> blocks) throws IOException {
        try (XMLSlideShow show = new XMLSlideShow()) {
            for (String block : blocks) {
                XSLFSlide slide = show.createSlide();
                XSLFTextBox box = slide.createTextBox();
                Dimension pageSize = show.getPageSize();
                box.setAnchor(new java.awt.geom.Rectangle2D.Double(40, 40, pageSize.width - 80, pageSize.height - 80));
                XSLFTextParagraph para = box.addNewTextParagraph();
                XSLFTextRun run = para.addNewTextRun();
                run.setText(block);
                run.setFontSize(24.0);
                run.setFontFamily("Calibri");
                run.setFontColor(Color.WHITE);
            }
            try (OutputStream out = Files.newOutputStream(output,
                    StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING)) {
                show.write(out);
            }
        }
    }

    private List<String> splitLyrics(String text) {
        if (text == null || text.isBlank()) return List.of();
        return Arrays.stream(text.split("\n\n+"))
                .map(String::trim).filter(b -> !b.isBlank()).toList();
    }

    @FXML void onCancel() { close(); }
    private void close() { ((Stage) btnCancel.getScene().getWindow()).close(); }
}
