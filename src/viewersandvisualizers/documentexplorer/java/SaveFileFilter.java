/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package viewersandvisualizers.documentexplorer.java;

import javax.swing.filechooser.FileFilter;
import java.io.File;

/**
 * Customizable File Filter used by the Save dialog.
 */
public class SaveFileFilter extends FileFilter {

    public SaveFileFilter(String extension, String description) {
        assert extension != null && !"".equals(extension) : "Null or Empty FileFilter extension.";

        mExtension = extension;
        mDescription = description;
    }

    /**
     * Returns true if the passed file matches the filter, false if the file
     * should be filtered.
     */
    public boolean accept(File f) {
        if (f.isDirectory()) {
            return true;
        }

        return f.getName().endsWith(mExtension);
    }

    /**
     * The description of this filter as displayed to the user in the filter
     * box.
     */
    public String getDescription() {
        return mDescription;
    }
    String mExtension;
    String mDescription;
}
