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
package com.aspose.examples;

import com.aspose.utils.AsposeConstants;
import com.aspose.utils.AsposeJavaAPI;
import com.aspose.utils.AsposeMavenProjectManager;
import com.aspose.utils.FormatExamples;
import com.aspose.utils.execution.ModalTaskImpl;
import com.intellij.openapi.progress.ProgressIndicator;
import com.intellij.openapi.progress.ProgressManager;
import com.intellij.openapi.ui.ComboBox;
import com.intellij.ui.components.JBScrollPane;
import com.intellij.ui.treeStructure.Tree;

import javax.swing.*;
import javax.swing.tree.DefaultTreeModel;
import javax.swing.tree.TreePath;
import java.awt.*;
import java.io.File;
import java.util.*;

/**
 * Created by Adeel Ilyas on 8/17/2015.
 */

public final class AsposeExamplePanel extends JPanel {

    AsposeExampleDialog dialog;

    /**
     * Creates new form AsposeExamplePanel
     */
    public AsposeExamplePanel(AsposeExampleDialog dialog) {
        initComponents();
        initComponentsUser();
        this.dialog = dialog;

    }

    private void initComponentsUser() {
        CustomMutableTreeNode top = new CustomMutableTreeNode("");
        read();
        validateDialog();
    }

    @Override
    public String getName() {
        return AsposeConstants.API_NAME + " for Java API and Examples";
    }

    private void initComponents() {
        ResourceBundle bundle = ResourceBundle.getBundle("Bundle");
        jPanel1 = new JPanel();
        jLabel2 = new JLabel();
        componentSelection = new ComboBox();

        jLabel1 = new JLabel();
        jLabelMessage = new JLabel();
        jLabelMessage.setOpaque(true);
        jScrollPane1 = new JBScrollPane();

        examplesTree = new Tree();

        jPanel1.setBackground(new Color(255, 255, 255));
        jPanel1.setBorder(BorderFactory.createEtchedBorder());
        jPanel1.setForeground(new Color(255, 255, 255));

        jLabel2.setIcon(icon); // NOI18N
        jLabel2.setText("");


        jLabel2.setHorizontalAlignment(SwingConstants.CENTER);
        jLabel2.setDoubleBuffered(true);
        jLabel2.setOpaque(true);
        jLabel2.addComponentListener(new java.awt.event.ComponentAdapter() {
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
            public void propertyChange(java.beans.PropertyChangeEvent evt) {
                componentSelection_Propertychanged(evt);
            }
        });
        jLabel1.setText(bundle.getString("AsposeExample.jLabel1_text"));
        jLabelMessage.setText("");
        examplesTree.addMouseListener(new java.awt.event.MouseAdapter() {
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

    //=========================================================================
    private void jLabel2ComponentResized(java.awt.event.ComponentEvent evt) {
        int labelwidth = jLabel2.getWidth();
        int labelheight = jLabel2.getHeight();
        Image img = icon.getImage();
        jLabel2.setIcon(new ImageIcon(img.getScaledInstance(labelwidth, labelheight, Image.SCALE_FAST)));
    }

    public String getSelectedProjectRootPath() {
        return AsposeMavenProjectManager.getInstance().getProjectHandle().getBasePath();
    }

    //=========================================================================
    void read() {
        AsposeConstants.println(" === New File Visual Panel.read() === " + AsposeMavenProjectManager.getInstance().getProjectHandle().getBaseDir().getName());
        retrieveAPIDependency();
        retrieveAPIExamples();
    }

    //=========================================================================
    private boolean retrieveAPIDependency() {
        getComponentSelection().removeAllItems();
        String versionNo = AsposeMavenProjectManager.getInstance().getDependencyVersionFromPOM(AsposeConstants.API_MAVEN_DEPENDENCY, AsposeMavenProjectManager.getInstance().getProjectHandle());
        if (versionNo == null) {
            getComponentSelection().addItem(AsposeConstants.API_DEPENDENCY_NOT_FOUND);
        } else {
            getComponentSelection().addItem(versionNo);
        }
        return true;
    }

    //=========================================================================
    private void retrieveAPIExamples() {
        final String item = (String) getComponentSelection().getSelectedItem();
        CustomMutableTreeNode top = new CustomMutableTreeNode("");
        DefaultTreeModel model = (DefaultTreeModel) getExamplesTree().getModel();
        model.setRoot(top);
        model.reload(top);
        System.out.println("ITEM: " + item);
        if (item != null && !item.equals(AsposeConstants.API_DEPENDENCY_NOT_FOUND)) {

            AsposeExampleCallback callback = new AsposeExampleCallback(this, top);
            final ModalTaskImpl modalTask = new ModalTaskImpl(AsposeMavenProjectManager.getInstance().getProjectHandle(), callback, "Please wait...");
            ProgressManager.getInstance().run(modalTask);
            top.setTopTreeNodeText(AsposeConstants.API_NAME);
            model.setRoot(top);
            model.reload(top);
            getExamplesTree().expandPath(new TreePath(top.getPath()));

        }

    }

//=========================================================================

    @Override
    public void validate() {
        AsposeConstants.println("AsposeExamplePanel validate called..");
    }

    //=========================================================================
    public boolean validateDialog() {
        if (isExampleSelected()) {
            if (dialog != null)
                dialog.updateControls(true);
            clearMessage();
            return true;
        }
        final String item = (String) getComponentSelection().getSelectedItem();
        if (item == null || item.equals(AsposeConstants.API_DEPENDENCY_NOT_FOUND)) {
            if (dialog != null)
                dialog.updateControls(false);
            diplayMessage("Please first add maven dependency of " + AsposeConstants.API_NAME + " for java API", true);
            return false;
        } else if (!isExampleSelected()) {
            if (dialog != null)
                dialog.updateControls(false);
            diplayMessage(AsposeConstants.ASPOSE_SELECT_EXAMPLE, true);
            return false;
        }
        if (dialog != null)
            dialog.updateControls(true);
        clearMessage();
        return true;
    }

    //=========================================================================

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
            return false;
        }
        return true;
    }

    //=========================================================================
    public void diplayMessage(String message, boolean error) {

        if (error) {
            jLabelMessage.setForeground(Color.RED);
        } else {
            jLabelMessage.setForeground(Color.GREEN);
        }
        jLabelMessage.setText(message);
    }

    //=========================================================================
    private void clearMessage() {
        jLabelMessage.setText("");

    }

    //=========================================================================
    public int showMessage(String title, String message, int buttons, int icon) {
        int result = JOptionPane.showConfirmDialog(null, message, title, buttons, icon);
        return result;
    }


    //====================================================================
    public void populateExamplesTree(AsposeJavaAPI asposeComponent, CustomMutableTreeNode top, ProgressIndicator p)

    {
        String examplesFullPath = asposeComponent.getLocalRepositoryPath() + File.separator + AsposeConstants.SOURCE_API_EXAMPLES_LOCATION;
        File directory = new File(examplesFullPath);
        AsposeConstants.println(examplesFullPath+ "exists?"+directory.exists());
        getExamplesTree().removeAll();
        top.setExPath(examplesFullPath);
        Queue<Object[]> queue = new LinkedList<>();
        queue.add(new Object[]{null, directory});

        while (!queue.isEmpty()) {
            Object[] _entry = queue.remove();
            File childFile = ((File) _entry[1]);
            CustomMutableTreeNode parentItem = ((CustomMutableTreeNode) _entry[0]);
            if (childFile.isDirectory()) {
                if (parentItem != null) {
                    CustomMutableTreeNode child = new CustomMutableTreeNode(FormatExamples.formatTitle(childFile.getName()));
                    child.setExPath(childFile.getAbsolutePath());
                    child.setFolder(true);
                    parentItem.add(child);
                    parentItem = child;
                } else {
                    parentItem = top;
                }
                for (File f : childFile.listFiles()) {
                    String fileName=f.getName().toLowerCase();
                    queue.add(new Object[]{parentItem, f});
                }
            } else if (childFile.isFile()) {

                    String title = FormatExamples.formatTitle(childFile.getName());
                    CustomMutableTreeNode child = new CustomMutableTreeNode(title);
                    child.setFolder(false);
                    parentItem.add(child);

            }
        }

    }


    private void componentSelection_Propertychanged(java.beans.PropertyChangeEvent evt) {

    }

    //=========================================================================
    private void examplesTree_clicked(java.awt.event.MouseEvent evt)
    {
        // TODO add your handling code here:
        TreePath path = getExamplesTree().getSelectionPath();

        validateDialog();
    }

    // Variables declaration
    private JComboBox componentSelection;
    private JTree examplesTree;
    private JLabel jLabel1;
    private JLabel jLabel2;
    private JLabel jLabelMessage;
    private JPanel jPanel1;
    private JScrollPane jScrollPane1;
    private ImageIcon icon = new ImageIcon(getClass().getResource("/resources/long_bannerIntelliJ.png"));
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