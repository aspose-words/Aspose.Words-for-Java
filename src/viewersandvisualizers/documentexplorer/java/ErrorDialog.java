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
import javax.swing.*;
import java.awt.event.*;

/**
 * Provides information to the user if an an unexpected exception occurs.
 */
public class ErrorDialog {

    public ErrorDialog(Exception ex) {

        errorDialog = new ErrorDialogForm();
        errorDialog.setModal(true);

        Toolkit toolkit = Toolkit.getDefaultToolkit();
        Dimension screenSize = toolkit.getScreenSize();

        // Calculate the frame location
        int x = (screenSize.width - errorDialog.getWidth()) / 2;
        int y = (screenSize.height - errorDialog.getHeight()) / 2;

        // Set the new frame location
        errorDialog.setLocation(x, y);

        errorDialog.closeButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onOK();
            }
        });

        errorDialog.addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent e) {
                onOK();
            }
        });

        errorDialog.setTitle(Globals.UNEXPECTED_EXCEPTION_DIALOG_TITLE);

        errorDialog.messageText.setText("\r\n" + ex.toString() + "\r\n");
        errorDialog.messageText.setSelectionStart(errorDialog.messageText.getText().length());

        errorDialog.setVisible(true);
    }

    private void onOK() {
        errorDialog.dispose();
    }
    ErrorDialogForm errorDialog;
}