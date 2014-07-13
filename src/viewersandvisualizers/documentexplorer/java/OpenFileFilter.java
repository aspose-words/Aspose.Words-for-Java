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
 * Customizable File Filter used by Open dialog.
 */
public class OpenFileFilter extends FileFilter {

    public OpenFileFilter(String[] extensions, String description) {
        assert extensions != null && extensions.length > 0 : "Null or Empty OpenFileFilter extensions array.";
        for (String ext : extensions) {
            assert ext != null && !"".equals(ext) : "Null or Empty OpenFileFilter extension.";
        }

        mExtensions = extensions;
        mDescription = description;
    }

    public boolean accept(File f) {
        if (f.isDirectory()) {
            return true;
        }

        for (String ext : mExtensions) {
            if (f.getName().endsWith(ext)) {
                return true;
            }
        }

        return false;
    }

    // The description of this filter
    public String getDescription() {
        return mDescription;
    }
    String[] mExtensions;
    String mDescription;
}
