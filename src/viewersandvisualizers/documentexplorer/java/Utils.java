/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package viewersandvisualizers.documentexplorer.java;

import javax.swing.*;
import java.io.File;

public class Utils {

    private Utils() {
    }

    /*
     * Gets the file extension of the passed file.
     */
    public static String getExtension(String s) {
        String ext = "";
        int i = s.lastIndexOf('.');

        if (i > 0 && i < s.length() - 1) {
            ext = s.substring(i + 1).toLowerCase();
        }

        return ext;
    }

    /*
     * Changes the extension of a file.
     *
     * Note: The extension should include a dot.
     */
    public static File setExtension(File f, String ext) {
        String name = f.getName();
        String newName;

        assert !"".equals(name) : "Empty file name.";

        // Don't change if the new extension is the same as the original.
        if (name.endsWith(ext)) {
            return f;
        }

        int lastIndexOfDot = name.lastIndexOf('.');

        if (lastIndexOfDot < 0) // File name without any extension.
        {
            newName = name + ext;
        } else // Change the existing extension.
        {
            newName = name.substring(0, lastIndexOfDot) + ext;
        }

        return new File(f.getParent(), newName);
    }

    /**
     * Returns an ImageIcon, or null if the path was invalid.
     */
    public static ImageIcon createImageIcon(String path) {
        java.net.URL imgURL = MainForm.class.getResource(path);
        if (imgURL != null) {
            return new ImageIcon(imgURL);
        } else {
            return null;
        }
    }
}