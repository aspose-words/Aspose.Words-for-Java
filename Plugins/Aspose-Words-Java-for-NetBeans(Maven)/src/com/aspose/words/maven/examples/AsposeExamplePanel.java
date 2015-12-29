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
package com.aspose.words.maven.examples;

import com.aspose.words.maven.utils.AbstractTask;
import com.aspose.words.maven.utils.AsposeConstants;
import com.aspose.words.maven.utils.AsposeMavenProjectManager;
import com.aspose.words.maven.utils.AsposeWordsJavaAPI;
import javax.swing.*;
import javax.swing.tree.TreePath;
import java.awt.*;
import javax.swing.tree.DefaultTreeModel;
import org.netbeans.api.progress.BaseProgressUtils;
import org.openide.WizardDescriptor;
import org.openide.util.Exceptions;
import org.openide.util.NbBundle;

/**
 * Created by Adeel Ilyas on 12/16/2015.
 */
public final class AsposeExamplePanel extends JPanel {

    AsposeExampleWizardPanel panel;

    /**
     * Creates new form AsposeExamplePanel
     *
     * @param panel
     */
    public AsposeExamplePanel(AsposeExampleWizardPanel panel) {
        initComponents();
        initComponentsUser();
        this.panel = panel;

    }

    private void initComponentsUser() {

        CustomMutableTreeNode top = new CustomMutableTreeNode("");

        DefaultTreeModel model = (DefaultTreeModel) getExamplesTree().getModel();
        model.setRoot(top);
        model.reload(top);

        validateDialog();
    }

    @Override
    public String getName() {
        return AsposeConstants.API_NAME + " for Java API - Code Examples";
    }

    private void initComponents() {

        jPanel1 = new JPanel();
        jLabel2 = new JLabel();
        componentSelection = new JComboBox();

        jLabel1 = new JLabel();
        jLabelMessage = new JLabel();
        jLabelMessage.setOpaque(true);
        jScrollPane1 = new JScrollPane();

        examplesTree = new JTree();

        jPanel1.setBackground(new Color(255, 255, 255));
        jPanel1.setBorder(BorderFactory.createEtchedBorder());
        jPanel1.setForeground(new Color(255, 255, 255));

        jLabel2.setIcon(icon); // NOI18N
        jLabel2.setText("");

        jLabel2.setHorizontalAlignment(SwingConstants.CENTER);
        jLabel2.setDoubleBuffered(true);
        jLabel2.setOpaque(true);
        jLabel2.addComponentListener(new java.awt.event.ComponentAdapter() {
            @Override
            public void componentResized(java.awt.event.ComponentEvent evt) {
                jLabel2ComponentResized(evt);
            }
        });

        GroupLayout jPanel1Layout = new GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
                jPanel1Layout.createParallelGroup()
                .addComponent(jLabel2, GroupLayout.PREFERRED_SIZE, 390, Short.MAX_VALUE)
        );
        jPanel1Layout.setVerticalGroup(
                jPanel1Layout.createParallelGroup()
                .addGroup(GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                        .addComponent(jLabel2)
                        .addGap(0, 0, Short.MAX_VALUE))
        );
        componentSelection.setModel(new DefaultComboBoxModel());

        componentSelection.addPropertyChangeListener(new java.beans.PropertyChangeListener() {
            @Override
            public void propertyChange(java.beans.PropertyChangeEvent evt) {
                componentSelection_Propertychanged(evt);
            }
        });
        jLabel1.setText(NbBundle.getMessage(AsposeExamplePanel.class, "AsposeExample.jLabel1_text"));
        jLabelMessage.setText("");
        examplesTree.addMouseListener(new java.awt.event.MouseAdapter() {
            @Override
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                examplesTree_clicked(evt);
            }
        });
        jScrollPane1.setViewportView(examplesTree);

        GroupLayout layout = new GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
                layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                .addComponent(jScrollPane1)
                .addGroup(layout.createSequentialGroup()
                        .addComponent(jLabel1)
                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(componentSelection, GroupLayout.PREFERRED_SIZE, 198, GroupLayout.PREFERRED_SIZE))
                .addComponent(jLabelMessage, GroupLayout.PREFERRED_SIZE, 361, GroupLayout.PREFERRED_SIZE)
                .addComponent(jPanel1, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
                layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                .addGroup(layout.createSequentialGroup()
                        .addComponent(jPanel1, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                .addComponent(componentSelection, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                                .addComponent(jLabel1, GroupLayout.PREFERRED_SIZE, 23, GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jScrollPane1, GroupLayout.PREFERRED_SIZE, 229, GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jLabelMessage, GroupLayout.PREFERRED_SIZE, 20, GroupLayout.PREFERRED_SIZE))
        );
    }


    private void jLabel2ComponentResized(java.awt.event.ComponentEvent evt) {
        int labelwidth = jLabel2.getWidth();
        int labelheight = jLabel2.getHeight();
        Image img = icon.getImage();
        jLabel2.setIcon(new ImageIcon(img.getScaledInstance(labelwidth, labelheight, Image.SCALE_FAST)));
    }

    /**
     *
     * @return
     */
    public String getSelectedProjectRootPath() {
        return AsposeMavenProjectManager.getInstance().getProjectDir().getPath();
    }

    void read() {
        retrieveAPIDependency();
        retrieveAPIExamples();
    }

    /**
     *
     * @return
     */
    private boolean retrieveAPIDependency() {
        getComponentSelection().removeAllItems();
        String versionNo = AsposeMavenProjectManager.getInstance().getDependencyVersionFromPOM(AsposeConstants.API_MAVEN_DEPENDENCY);
        if (versionNo == null) {
            getComponentSelection().addItem(AsposeConstants.API_DEPENDENCY_NOT_FOUND);
        } else {
            getComponentSelection().addItem(versionNo);
        }
        return true;
    }

    private void retrieveAPIExamples() {

        final String item = (String) getComponentSelection().getSelectedItem();

        if (item != null && !item.equals(AsposeConstants.API_DEPENDENCY_NOT_FOUND)) {

            // Downloading Aspose API mvn based examples
            AbstractTask downloadExamples = AsposeMavenProjectManager.getInstance().createDownloadExamplesTask(AsposeWordsJavaAPI.getInstance());
            // Execute the tasks
            BaseProgressUtils.showProgressDialogAndRun(downloadExamples, NbBundle.getMessage(AsposeExamplePanel.class, "AsposeManager.updateExamplesMessage"));

            // Populating Aspose API mvn based examples
            Runnable popuplateExamples = AsposeMavenProjectManager.getInstance().populateExamplesTask(AsposeWordsJavaAPI.getInstance(), this);
            // Execute the tasks
            BaseProgressUtils.showProgressDialogAndRun(popuplateExamples, NbBundle.getMessage(AsposeExamplePanel.class, "AsposeManager.populateExamplesMessage"));

            validateDialog();

        }

    }

    boolean valid(WizardDescriptor wizardDescriptor) {

        return validateDialog();
    }

    @Override
    public void validate() {

    }

    /**
     *
     * @return
     */
    public boolean validateDialog() {
        if (isExampleSelected()) {
            clearMessage();
            return true;
        }
        final String item = (String) getComponentSelection().getSelectedItem();
        if (item == null || item.equals(AsposeConstants.API_DEPENDENCY_NOT_FOUND)) {
            diplayMessage("Please first add maven dependency of " + AsposeConstants.API_NAME + " for java API", true);
            return false;
        } else if (!isExampleSelected()) {
            diplayMessage(AsposeConstants.ASPOSE_SELECT_EXAMPLE, true);
            return false;
        }
        clearMessage();
        return true;
    }

    /**
     *
     * @return
     */
    private boolean isExampleSelected() {
        CustomMutableTreeNode comp = (CustomMutableTreeNode) getExamplesTree().getLastSelectedPathComponent();
        if (comp == null) {
            return false;
        }
        try {

            if (!comp.isFolder()) {
                return false;
            }
        } catch (Exception ex) {
            Exceptions.printStackTrace(ex);
            return false;
        }
        return true;
    }

    /**
     *
     * @param message
     * @param error
     */
    public void diplayMessage(String message, boolean error) {

        if (error) {
            jLabelMessage.setForeground(Color.RED);
        } else {
            jLabelMessage.setForeground(Color.GREEN);
        }
        jLabelMessage.setText(message);
    }

    private void clearMessage() {
        jLabelMessage.setText("");

    }

    /**
     *
     * @param title
     * @param message
     * @param buttons
     * @param icon
     * @return
     */
    public int showMessage(String title, String message, int buttons, int icon) {
        int result = JOptionPane.showConfirmDialog(null, message, title, buttons, icon);
        return result;
    }

    private void componentSelection_Propertychanged(java.beans.PropertyChangeEvent evt) {

    }

    private void examplesTree_clicked(java.awt.event.MouseEvent evt) {
        TreePath path = getExamplesTree().getSelectionPath();
        panel.fireChangeEvent();
    }

    // Variables declaration
    private JComboBox componentSelection;
    private JTree examplesTree;
    private JLabel jLabel1;
    private JLabel jLabel2;
    private JLabel jLabelMessage;
    private JPanel jPanel1;
    private JScrollPane jScrollPane1;
    private ImageIcon icon = new ImageIcon(getClass().getResource("/resources/long_banner.png"));
    // End of variables declaration

    /**
     * @return the examplesTree
     */
    public JTree getExamplesTree() {
        return examplesTree;
    }

    /**
     * @return the componentSelection
     */
    public JComboBox getComponentSelection() {
        return componentSelection;
    }
}
