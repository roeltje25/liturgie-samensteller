package nl.roeltje.liturgie.ui;

import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;
import nl.roeltje.liturgie.AppContext;

import java.io.File;
import java.io.IOException;
import java.nio.file.*;
import java.util.List;

/**
 * Import songs from existing PPTX archive.
 */
public class ImportPptxController implements DialogController {

    @FXML private TextField tfFolder;
    @FXML private TableView<ScanRow> table;
    @FXML private TableColumn<ScanRow, String> colFile, colDate, colSongs;
    @FXML private ProgressBar progressBar;
    @FXML private Label lblStatus;
    @FXML private Button btnBrowse, btnScan, btnImport, btnCancel;

    private final AppContext ctx = AppContext.get();

    public record ScanRow(String file, String date, String songs, Path path) {}

    @FXML
    public void initialize() {
        colFile.setCellValueFactory(d -> new javafx.beans.property.SimpleStringProperty(d.getValue().file()));
        colDate.setCellValueFactory(d -> new javafx.beans.property.SimpleStringProperty(d.getValue().date()));
        colSongs.setCellValueFactory(d -> new javafx.beans.property.SimpleStringProperty(d.getValue().songs()));
        progressBar.setVisible(false);
        tfFolder.setText(ctx.getSettings().getPptxArchivePath().toString());
        btnBrowse.setText(ctx.tr("button.browse"));
        btnScan.setText(ctx.tr("dialog.import_pptx.scan"));
        btnImport.setText(ctx.tr("dialog.import_pptx.import_excel"));
        btnCancel.setText(ctx.tr("button.cancel"));
    }

    @FXML void onBrowse() {
        DirectoryChooser dc = new DirectoryChooser();
        File dir = dc.showDialog(tfFolder.getScene().getWindow());
        if (dir != null) tfFolder.setText(dir.getAbsolutePath());
    }

    @FXML void onScan() {
        Path folder = Path.of(tfFolder.getText().trim());
        if (!Files.isDirectory(folder)) {
            lblStatus.setText(ctx.tr("dialog.import_pptx.folder_not_found", "path", folder));
            return;
        }
        table.getItems().clear();
        progressBar.setVisible(true);
        progressBar.setProgress(ProgressIndicator.INDETERMINATE_PROGRESS);
        lblStatus.setText(ctx.tr("dialog.import_pptx.scanning"));

        Task<List<ScanRow>> task = new Task<>() {
            @Override protected List<ScanRow> call() throws IOException {
                List<ScanRow> rows = new java.util.ArrayList<>();
                try (var stream = Files.list(folder)) {
                    List<Path> pptxFiles = stream
                            .filter(p -> p.getFileName().toString().toLowerCase().endsWith(".pptx"))
                            .sorted().toList();
                    for (Path p : pptxFiles) {
                        var slides = ctx.getPptxService().getSlidesInfo(p);
                        rows.add(new ScanRow(
                                p.getFileName().toString(),
                                ctx.tr("dialog.import_pptx.no_date"),
                                ctx.tr("dialog.import_pptx.songs_found", "count", slides.size()),
                                p));
                    }
                }
                return rows;
            }
        };
        task.setOnSucceeded(e -> {
            table.getItems().addAll(task.getValue());
            progressBar.setVisible(false);
            lblStatus.setText(ctx.tr("dialog.import_pptx.scan_complete", "count", table.getItems().size()));
        });
        task.setOnFailed(e -> {
            progressBar.setVisible(false);
            lblStatus.setText(ctx.tr("dialog.import_pptx.scan_error", "error", task.getException().getMessage()));
        });
        new Thread(task, "scan-pptx").start();
    }

    @FXML void onImportExcel() {
        Path xl = ctx.getSettings().getExcelRegisterPath();
        if (xl == null) { lblStatus.setText(ctx.tr("dialog.import_pptx.no_excel")); return; }
        lblStatus.setText(ctx.tr("dialog.import_pptx.import_complete", "count", table.getItems().size()));
    }

    @FXML void onCancel() { ((Stage) btnCancel.getScene().getWindow()).close(); }
}
