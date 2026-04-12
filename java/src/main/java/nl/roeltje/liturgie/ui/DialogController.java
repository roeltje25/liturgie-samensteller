package nl.roeltje.liturgie.ui;

/**
 * Marker interface for dialog controllers that can receive an argument
 * and a reference to the main controller before the dialog is shown.
 */
public interface DialogController {
    /** Optional string argument (section ID, file path, etc.) */
    default void setArg(String arg) {}
    /** Reference to the main controller so dialogs can add items. */
    default void setMainController(MainController main) {}
}
