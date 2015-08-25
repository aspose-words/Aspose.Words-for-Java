/*
 * The MIT License (MIT)
 *
 * Copyright (c) 1998-2015 Aspose Pty Ltd.
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */
package com.aspose.wizards.maven;


import com.intellij.ide.util.projectWizard.ModuleWizardStep;
import com.intellij.ide.wizard.CommitStepException;
import com.intellij.openapi.Disposable;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.util.ResourceBundle;

/**
 * @author Adeel Ilyas
 */

public class AsposeIntroWizardStep extends ModuleWizardStep implements Disposable {

    @Override
    public void dispose() {
    }

    private JLabel jLabelAsposeIntro;
    private JComponent myMainPanel;

    @Override
    public JComponent getPreferredFocusedComponent() {
        return myMainPanel;
    }

    @Override
    public void onWizardFinished() throws CommitStepException {

    }

    @Override
    public JComponent getComponent() {

        if (myMainPanel == null) {
            myMainPanel = new JPanel();
            {
                ResourceBundle bundle = ResourceBundle.getBundle("Bundle");
                myMainPanel.setBorder(new TitledBorder(bundle.getString("AsposeWizardPanel.myMainPanel.border.title")));
                myMainPanel.setPreferredSize(new Dimension(333, 364));


                jLabelAsposeIntro = new JLabel();
                jLabelAsposeIntro.setText(bundle.getString("AsposeWizardPanel.myMainPanel.description"));
                Font labelFont = jLabelAsposeIntro.getFont();

                jLabelAsposeIntro.setFont(new Font(labelFont.getName(), Font.PLAIN, 14));

                GroupLayout jPanel4Layout = new GroupLayout(myMainPanel);
                myMainPanel.setLayout(jPanel4Layout);
                jPanel4Layout.setHorizontalGroup(
                        jPanel4Layout.createParallelGroup()
                                .addGroup(jPanel4Layout.createSequentialGroup()
                                        .addComponent(jLabelAsposeIntro)
                                        .addGap(0, 0, Short.MAX_VALUE))
                );
                jPanel4Layout.setVerticalGroup(
                        jPanel4Layout.createParallelGroup()
                                .addGroup(jPanel4Layout.createSequentialGroup()
                                        .addComponent(jLabelAsposeIntro)
                                        .addContainerGap(0, Short.MAX_VALUE))
                );
            }
        }
        return myMainPanel;
    }

    @Override
    public void updateDataModel() {

    }


}