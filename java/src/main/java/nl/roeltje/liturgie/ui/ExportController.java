package nl.roeltje.liturgie.ui;

import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.Stage;
import nl.roeltje.liturgie.AppContext;
import nl.roeltje.liturgie.services.ExportService;

import java.nio.file.*;

/**
 * Export dialog controller.
 */
public class ExportController implements DialogController {

    @FXML private TextField tfFilename;
    @FXML private CheckBox cbPptx, cbPdf, cbTxt, cbExcel;
    @FXML private CheckBox cbOpenAfter;
    @FXML private ProgressBar progressBar;
    @FXML private Label lblProgress;
    @FXML private Button btnExport, btnCancel;

    private final AppContext ctx = AppContext.get();

    @FXML
    public void initialize() {
        if (ctx.getCurrentLiturgy() != null) {
            tfFilename.setText(ctx.getExportService().getDefaultFilename(ctx.getCurrentLiturgy()));
        }
        cbPptx.setSelected(true);
        cbPdf.setSelected(true);
        cbTxt.setSelected(true);
        cbExcel.setSelected(ctx.getSettings().getExcelRegisterPath() != null);
        progressBar.setVisible(false);

        btnExport.setText(ctx.tr("button.export"));
        btnCancel.setText(ctx.tr("button.cancel"));
    }

    @FXML void onExport() {
        if (ctx.getCurrentLiturgy() == null) return;
        btnExport.setDisable(true);
        progressBar.setVisible(true);
        progressBar.setProgress(ProgressIndicator.INDETERMINATE_PROGRESS);
        lblProgress.setText(ctx.tr("status.exporting"));

        Task<ExportService.ExportResult> task = new Task<>() {
            @Override protected ExportService.ExportResult call() {
                return ctx.getExportService().exportAll(
                        ctx.getCurrentLiturgy(),
                        cbPptx.isSelected(),
                        cbPdf.isSelected(),
                        cbTxt.isSelected(),
                        cbExcel.isSelected());
            }
        };
        task.setOnSucceeded(e -> {
            ExportService.ExportResult result = task.getValue();
            progressBar.setProgress(1.0);
            if (result.errors().isEmpty()) {
                lblProgress.setText(ctx.tr("dialog.export.success"));
                if (cbOpenAfter.isSelected() && result.files().containsKey("pptx")) {
                    openFile(result.files().get("pptx"));
                }
            } else {
                lblProgress.setText("Errors: " + String.join("; ", result.errors()));
            }
            btnExport.setDisable(false);
        });
        task.setOnFailed(e -> {
            lblProgress.setText("Export failed: " + task.getException().getMessage());
            btnExport.setDisable(false);
        });
        new Thread(task, "export-task").start();
    }

    @FXML void onCancel() { ((Stage) btnCancel.getScene().getWindow()).close(); }

    private void openFile(Path file) {
        try {
            java.awt.Desktop.getDesktop().open(file.toFile());
        } catch (Exception e) {
            // Not critical
        }
    }
}
