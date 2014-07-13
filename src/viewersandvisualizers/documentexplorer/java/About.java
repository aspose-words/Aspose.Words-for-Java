/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package viewersandvisualizers.documentexplorer.java;

import java.awt.Dimension;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

/**
 * Shows an About form for the DocumentExplorer application.
 */
public class About {

    public About() {
        aboutForm = new AboutForm();
        aboutForm.setModal(true);

        Toolkit toolkit = Toolkit.getDefaultToolkit();
        Dimension screenSize = toolkit.getScreenSize();

        // Calculate the frame location
        int x = (screenSize.width - aboutForm.getWidth()) / 2;
        int y = (screenSize.height - aboutForm.getHeight()) / 2;

        // Set the new frame location
        aboutForm.setLocation(x, y);

        aboutForm.closeButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onOK();
            }
        });

        aboutForm.setVisible(true);
    }

    private void onOK() {
        aboutForm.dispose();
    }
    AboutForm aboutForm;
}
