/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package viewersandvisualizers.documentexplorer.java;

import com.aspose.words.FileFormatUtil;
import com.aspose.words.SaveFormat;

import javax.swing.*;
import java.io.File;
import java.text.MessageFormat;

/**
 * Provides static methods for actions that involve dialogs.
 */
public class Dialogs {

    /**
     * Stores the last accessed file path, so that the next file dialog session
     * should start from that directory.
     */
    private static String mDocumentPath = "";
    /**
     * Stores the name of the currently loaded document. Name is stored without
     * extension to be used when saving in different file format.
     */
    private static String mDocumentName;
    /**
     * Stores an instance of a file chooser dialog which allows the user to
     * choose a document to open from disk.
     */
    private static JFileChooser mOpenDialog;
    /**
     * Stores an instance of a file chooser dialog which allows the user to save
     * a document to disk in a given format.
     */
    private static JFileChooser mSaveDialog;

    /**
     * This class is purely static.
     */
    private Dialogs() {
    }

    // Static ctor which sets up the open and save dialogs
    static {
        mOpenDialog = new JFileChooser();
        mOpenDialog.setAcceptAllFileFilterUsed(false);
        mOpenDialog.setFileFilter(Globals.OPEN_FILE_FILTER_DOC_FORMAT);
        mOpenDialog.setFileFilter(Globals.OPEN_FILE_FILTER_DOCX_FORMAT);
        mOpenDialog.setFileFilter(Globals.OPEN_FILE_FILTER_XML_FORMAT);
        mOpenDialog.setFileFilter(Globals.OPEN_FILE_FILTER_RTF_FORMAT);
        mOpenDialog.setFileFilter(Globals.OPEN_FILE_FILTER_ODT_FORMAT);
        mOpenDialog.setFileFilter(Globals.OPEN_FILE_FILTER_HTML_FORMAT);
        mOpenDialog.setFileFilter(Globals.OPEN_FILE_FILTER_ALL_SUPPORTED_FORMATS); // This is last so it will appear by default.
        mOpenDialog.setMultiSelectionEnabled(false);
        mOpenDialog.setFileSelectionMode(JFileChooser.FILES_ONLY);
        mOpenDialog.setDialogTitle(Globals.OPEN_DOCUMENT_DIALOG_TITLE);

        mSaveDialog = new JFileChooser();
        mSaveDialog.setMultiSelectionEnabled(false);
        mSaveDialog.setFileSelectionMode(JFileChooser.FILES_ONLY);
        mSaveDialog.setDialogTitle(Globals.SAVE_DOCUMENT_DIALOG_TITLE);
        mSaveDialog.setAcceptAllFileFilterUsed(false);

        mSaveDialog.addChoosableFileFilter(Globals.SAVE_FILE_FILTER_DOCX);
        mSaveDialog.addChoosableFileFilter(Globals.SAVE_FILE_FILTER_DOCM);
        mSaveDialog.addChoosableFileFilter(Globals.SAVE_FILE_FILTER_PDF);
        mSaveDialog.addChoosableFileFilter(Globals.SAVE_FILE_FILTER_XPS);
        mSaveDialog.addChoosableFileFilter(Globals.SAVE_FILE_FILTER_PDT);
        mSaveDialog.addChoosableFileFilter(Globals.SAVE_FILE_FILTER_HTML);
        mSaveDialog.addChoosableFileFilter(Globals.SAVE_FILE_FILTER_MHT);
        mSaveDialog.addChoosableFileFilter(Globals.SAVE_FILE_FILTER_RTF);
        mSaveDialog.addChoosableFileFilter(Globals.SAVE_FILE_FILTER_XML);
        mSaveDialog.addChoosableFileFilter(Globals.SAVE_FILE_FILTER_FOPC);
        mSaveDialog.addChoosableFileFilter(Globals.SAVE_FILE_FILTER_TXT);
        mSaveDialog.addChoosableFileFilter(Globals.SAVE_FILE_FILTER_EPUB);
        mSaveDialog.addChoosableFileFilter(Globals.SAVE_FILE_FILTER_SWF);
        mSaveDialog.addChoosableFileFilter(Globals.SAVE_FILE_FILTER_XAML);
        mSaveDialog.addChoosableFileFilter(Globals.SAVE_FILE_FILTER_DOC); // Add this last to make it the default format.

        // Corrects the behaviour of the "save as" dialog to be more like MS Word. Changing file filter
        // will automatically preserve file name and change to the appropriate extension.
        mSaveDialog.addPropertyChangeListener(new SaveDialogChangeListener(mSaveDialog));
    }

    /**
     * Opens a chooser dialog and allows the user to pick a document to open.
     *
     * @return The file name of the selected document or an empty string if the
     * dialog was canceled.
     */
    public static String openDocument() {
        mOpenDialog.setCurrentDirectory(new File(mDocumentPath));

        if (mOpenDialog.showOpenDialog(Globals.mMainForm) == JFileChooser.APPROVE_OPTION) {
            File file = mOpenDialog.getSelectedFile();
            String fileName = file.getAbsolutePath();
            if (file.exists()) {
                mDocumentPath = file.getParent();
                mDocumentName = file.getName();
                return fileName;
            } else {
                JOptionPane.showMessageDialog(Globals.mMainForm,
                        MessageFormat.format("File \"{0}\" doesn't exist.", fileName),
                        Globals.APPLICATION_TITLE, JOptionPane.ERROR_MESSAGE);
                return "";
            }
        } else {
            return "";
        }
    }

    /**
     * Opens a chooser dialog and allows the user to pick a path on the disk and
     * format to save a document to.
     *
     * @return The path to save the document to or an empty string if the dialog
     * was canceled.
     */
    public static String saveDocument() {
        mSaveDialog.setCurrentDirectory(new File(mDocumentPath));
        File currentDocument = new File(mDocumentName);
        mSaveDialog.setSelectedFile(Utils.setExtension(currentDocument, getCurrentSaveFileFilterExtension()));

        if (mSaveDialog.showSaveDialog(Globals.mMainForm) == JFileChooser.APPROVE_OPTION) {
            File file = mSaveDialog.getSelectedFile();

            // The format to save to is inferred from the extension of the file name. Check if this matches a valid Aspose.Words save format.
            // If the extension is unsupported or missing then add the extension of the current file filter and save in that format.
            if (FileFormatUtil.extensionToSaveFormat(Utils.getExtension(file.getName())) == SaveFormat.UNKNOWN) {
                String ext = getCurrentSaveFileFilterExtension();
                file = Utils.setExtension(file, ext);
            }

            // If the file exists open a dialog asking if the existing file should be overridden.
            if (file.exists()) {
                String fileName = file.getAbsolutePath();
                if (JOptionPane.showConfirmDialog(Globals.mMainForm,
                        java.text.MessageFormat.format("File \"{0}\" already exists. Would you like to overwrite it ?", fileName),
                        Globals.APPLICATION_TITLE, JOptionPane.YES_NO_OPTION) == JOptionPane.NO_OPTION) {
                    return "";
                }
            }

            mDocumentPath = file.getParent();
            return file.getAbsolutePath();
        } else {
            return "";
        }
    }

    /**
     * Returns the extension of the current file filter selected in the Save As
     * dialog.
     */
    private static String getCurrentSaveFileFilterExtension() {
        return ((SaveFileFilter) mSaveDialog.getFileFilter()).mExtension;
    }
}