package nl.roeltje.liturgie.ui;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.Stage;
import nl.roeltje.liturgie.AppContext;

/**
 * About dialog.
 */
public class AboutController implements DialogController {

    @FXML private Label lblTitle, lblVersion, lblDescription, lblJava, lblOs;
    @FXML private Button btnClose;

    private final AppContext ctx = AppContext.get();

    @FXML
    public void initialize() {
        lblTitle.setText(ctx.tr("app.title"));
        lblVersion.setText(ctx.tr("dialog.about.version", "version", AppContext.VERSION));
        lblDescription.setText(ctx.tr("dialog.about.description"));
        lblJava.setText("Java " + Runtime.version().feature() +
                " – " + System.getProperty("java.vm.name", ""));
        lblOs.setText(System.getProperty("os.name", "") + " " + System.getProperty("os.version", ""));
        btnClose.setText(ctx.tr("button.close"));
    }

    @FXML void onClose() { ((Stage) btnClose.getScene().getWindow()).close(); }
}
