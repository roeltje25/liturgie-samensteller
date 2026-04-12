package nl.roeltje.liturgie.ui;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import nl.roeltje.liturgie.AppContext;
import nl.roeltje.liturgie.models.Settings;

import java.io.File;

/**
 * Settings dialog controller.
 */
public class SettingsController implements DialogController {

    @FXML private TextField tfBaseFolder;
    @FXML private TextField tfSongsFolder, tfAlgemeenFolder, tfThemesFolder, tfOutputFolder;
    @FXML private TextField tfCollecteFile, tfStubTemplate, tfBibleTemplate, tfOutputPattern;
    @FXML private TextField tfExcelRegister, tfPptxArchive;
    @FXML private CheckBox cbSongCover;
    @FXML private TextField tfSongCoverFile;
    @FXML private ComboBox<String> cbLanguage;
    @FXML private TextField tfBibleFont;
    @FXML private Spinner<Integer> spBibleFontSize, spBibleCharsPerSlide;
    @FXML private CheckBox cbShowVerseNumbers;
    @FXML private TextField tfYouVersionKey;
    @FXML private Button btnSave, btnCancel;

    private final AppContext ctx = AppContext.get();

    @FXML
    public void initialize() {
        Settings s = ctx.getSettings();
        tfBaseFolder.setText(s.getBase_folder());
        tfSongsFolder.setText(s.getSongs_folder());
        tfAlgemeenFolder.setText(s.getAlgemeen_folder());
        tfThemesFolder.setText(s.getThemes_folder());
        tfOutputFolder.setText(s.getOutput_folder());
        tfCollecteFile.setText(s.getCollecte_filename());
        tfStubTemplate.setText(s.getStub_template_filename());
        tfBibleTemplate.setText(s.getBible_template_filename());
        tfOutputPattern.setText(s.getOutput_pattern());
        tfExcelRegister.setText(s.getExcel_register_path());
        tfPptxArchive.setText(s.getPptx_archive_folder());
        cbSongCover.setSelected(s.isSong_cover_enabled());
        tfSongCoverFile.setText(s.getSong_cover_filename());
        cbLanguage.getItems().addAll("nl", "en");
        cbLanguage.setValue(s.getLanguage());
        tfBibleFont.setText(s.getBible_font_name());
        spBibleFontSize.setValueFactory(new SpinnerValueFactory.IntegerSpinnerValueFactory(6, 72, s.getBible_font_size()));
        spBibleCharsPerSlide.setValueFactory(new SpinnerValueFactory.IntegerSpinnerValueFactory(100, 2000, s.getBible_chars_per_slide()));
        cbShowVerseNumbers.setSelected(s.isBible_show_verse_numbers());
        tfYouVersionKey.setText(s.getYouversion_api_key());
    }

    @FXML void onBrowseBase() {
        DirectoryChooser dc = new DirectoryChooser();
        dc.setTitle(ctx.tr("dialog.settings.base_folder"));
        File dir = dc.showDialog(tfBaseFolder.getScene().getWindow());
        if (dir != null) tfBaseFolder.setText(dir.getAbsolutePath());
    }

    @FXML void onBrowseExcel() {
        FileChooser fc = new FileChooser();
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel (*.xlsx)", "*.xlsx"));
        File file = fc.showOpenDialog(tfExcelRegister.getScene().getWindow());
        if (file != null) tfExcelRegister.setText(file.getAbsolutePath());
    }

    @FXML void onSave() {
        Settings s = ctx.getSettings();
        s.setBase_folder(tfBaseFolder.getText().trim());
        s.setSongs_folder(tfSongsFolder.getText().trim());
        s.setAlgemeen_folder(tfAlgemeenFolder.getText().trim());
        s.setThemes_folder(tfThemesFolder.getText().trim());
        s.setOutput_folder(tfOutputFolder.getText().trim());
        s.setCollecte_filename(tfCollecteFile.getText().trim());
        s.setStub_template_filename(tfStubTemplate.getText().trim());
        s.setBible_template_filename(tfBibleTemplate.getText().trim());
        s.setOutput_pattern(tfOutputPattern.getText().trim());
        s.setExcel_register_path(tfExcelRegister.getText().trim());
        s.setPptx_archive_folder(tfPptxArchive.getText().trim());
        s.setSong_cover_enabled(cbSongCover.isSelected());
        s.setSong_cover_filename(tfSongCoverFile.getText().trim());
        s.setLanguage(cbLanguage.getValue());
        s.setBible_font_name(tfBibleFont.getText().trim());
        s.setBible_font_size(spBibleFontSize.getValue());
        s.setBible_chars_per_slide(spBibleCharsPerSlide.getValue());
        s.setBible_show_verse_numbers(cbShowVerseNumbers.isSelected());
        s.setYouversion_api_key(tfYouVersionKey.getText().trim());
        ctx.updateSettings(s);
        ctx.getFolderScanner().refresh();
        close();
    }

    @FXML void onCancel() { close(); }

    private void close() { ((Stage) btnCancel.getScene().getWindow()).close(); }
}
