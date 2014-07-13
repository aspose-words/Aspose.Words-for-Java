/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package viewersandvisualizers.documentexplorer.java;

import javax.swing.*;
import java.beans.PropertyChangeListener;
import java.beans.PropertyChangeEvent;
import java.io.File;

/**
 * Corrects JFileChooser Save Dialog behavior to act more like how Microsoft
 * Word does. - Preserves the current file name in the edit box when a different
 * filter or directory is chosen. - Automatically changes the file extension on
 * the file name when a different filter is chosen.
 */
class SaveDialogChangeListener implements PropertyChangeListener {

    public SaveDialogChangeListener(JFileChooser chooser) {
        mChooser = chooser;
    }

    public void propertyChange(PropertyChangeEvent e) {
        String propertyName = e.getPropertyName();

        if (JFileChooser.SELECTED_FILE_CHANGED_PROPERTY.equals(propertyName)) {
            fileChanged(e);
        } else if (JFileChooser.FILE_FILTER_CHANGED_PROPERTY.equals(propertyName)) {
            filterChanged(e);
        } else if (JFileChooser.DIRECTORY_CHANGED_PROPERTY.equals(propertyName)) {
            directoryChanged(e);
        }

    }

    /**
     * Remembers the old and new file names used in the dialog.
     */
    private void fileChanged(PropertyChangeEvent e) {
        mNewFile = (File) e.getNewValue();
        mOldFile = (File) e.getOldValue();
    }

    /**
     * Changes a file extension when the user changes file filter.
     */
    private void filterChanged(PropertyChangeEvent e) {
        // When the user changes the filter used, JFileChooser deletes the "old" filename - we must correct this.
        if (mNewFile == null && mOldFile != null) {
            mNewFile = mOldFile;
        }

        // Change the file extension according to the new chosen filter.
        SaveFileFilter newFilter = (SaveFileFilter) e.getNewValue();
        mNewFile = Utils.setExtension(mNewFile, newFilter.mExtension);

        // Show this to the user.
        mChooser.setSelectedFile(mNewFile);
        mChooser.updateUI();
    }

    /**
     * Restores a file name deleted by JFileChooser when an user changes current
     * directory.
     */
    private void directoryChanged(PropertyChangeEvent e) {
        if (mNewFile == null && mOldFile != null) {
            mNewFile = new File(mOldFile.getName());

            mChooser.setSelectedFile(mNewFile);
            mChooser.updateUI();
        }
    }
    private static JFileChooser mChooser;
    private static File mNewFile;
    private static File mOldFile;
}
