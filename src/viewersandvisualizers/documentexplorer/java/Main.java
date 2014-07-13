/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package viewersandvisualizers.documentexplorer.java;

import com.aspose.words.Document;
import com.aspose.words.License;

import javax.swing.event.TreeExpansionEvent;
import javax.swing.event.TreeSelectionEvent;
import javax.swing.event.TreeSelectionListener;
import javax.swing.event.TreeWillExpandListener;
import javax.swing.tree.*;
import java.awt.*;
import java.awt.event.*;
import java.io.File;
import java.util.Enumeration;
import javax.swing.ImageIcon;
import javax.swing.JTree;

/**
 * The main form of the DocumentExplorer demo.
 * 
* DocumentExplorer allows to open DOC, DOT, DOCX, XML, WML, RTF, ODT, OTT,
 * HTML, XHTML and MHTML files using Aspose.Words.
 * 
* Once a document is opened, you can explore its object model in the tree. You
 * can also save the document into DOC, DOCX, ODF, EPUB, PDF, SWF, RTF, WordML,
 * HTML, MHTML and plain text formats.
 *
 */
public class Main implements TreeWillExpandListener, TreeSelectionListener, KeyListener {

    public Main() throws Exception {

        // Search for an Aspose.Words license in the application directory.
        // The File.Exists check is only needed in this demo so it will work
        // both when the license file is present as well as when it's missing.
        // In your real application you just need to call the SetLicense method.
        File licenseFile = new File(System.getProperty("user.dir") + "\\Aspose.Words.lic");
        if (licenseFile.exists()) {
            // This shows how to license Aspose.Words.
            // If you don't specify a license, Aspose.Words works in evaluation mode.
            License license = new License();
            license.setLicense(licenseFile.getAbsolutePath());
        }

        Globals.mMainForm = new MainForm();

        // Get the screen size
        Toolkit toolkit = Toolkit.getDefaultToolkit();
        Dimension screenSize = toolkit.getScreenSize();

        // Calculate the frame location
        int x = (screenSize.width - Globals.mMainForm.getWidth()) / 2;
        int y = (screenSize.height - Globals.mMainForm.getHeight()) / 2;

        // Set the new frame location
        Globals.mMainForm.setLocation(x, y);

        Globals.mMainForm.setTitle(Globals.APPLICATION_TITLE);
        Globals.mMainForm.setIconImage(Utils.createImageIcon("images/App.gif").getImage());

        Globals.mMainForm.addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent e) {
                onClose();
            }
        });

        Globals.mMainForm.toolOpenDocument.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent evt) {
                onOpen();
            }
        });

        Globals.mMainForm.toolSaveAs.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent evt) {
                onSaveAs();
            }
        });

        Globals.mMainForm.toolExpandAll.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent evt) {
                onExpandAll();
            }
        });

        Globals.mMainForm.toolCollapseAll.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent evt) {
                onCollapseAll();
            }
        });

        Globals.mMainForm.toolRemove.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent evt) {
                onRemove();
            }
        });

        Globals.mMainForm.menuOpen.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent evt) {
                onOpen();
            }
        });

        Globals.mMainForm.menuSaveAs.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent evt) {
                onSaveAs();
            }
        });
        Globals.mMainForm.menuExit.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent evt) {
                onClose();
            }
        });

        Globals.mMainForm.menuRemoveNode.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent evt) {
                onRemove();
            }
        });

        Globals.mMainForm.menuExpandAll.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent evt) {
                onExpandAll();
            }
        });

        Globals.mMainForm.menuCollapseAll.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent evt) {
                onCollapseAll();
            }
        });

        Globals.mMainForm.menuAbout.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent evt) {
                onAbout();
            }
        });

        Globals.mMainForm.setVisible(true);
    }

    private void onClose() {
        Globals.mMainForm.dispose();
    }

    /**
     * Opens a document with the name and format provided in a standard Save As
     * dialog.
     */
    private void onOpen() {

        try {
            String fileName = Dialogs.openDocument();
            if (!"".equals(fileName)) {
                Globals.mMainForm.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
                Globals.mDocument = new Document(fileName);

                Globals.mMainForm.setTitle(Globals.APPLICATION_TITLE + " - " + fileName);

                Globals.mRootNode = Item.createItem(Globals.mDocument).getTreeNode();
                Globals.mTreeModel = new DefaultTreeModel(Globals.mRootNode);
                Globals.mTree = new JTree(Globals.mTreeModel);
                Globals.mTree.setExpandsSelectedPaths(false);
                Globals.mTree.getSelectionModel().setSelectionMode(TreeSelectionModel.SINGLE_TREE_SELECTION);
                Globals.mTree.setCellRenderer(new OurCellRenderer());
                Globals.mTree.setShowsRootHandles(true);
                Globals.mTree.addTreeWillExpandListener(this);
                Globals.mTree.addTreeSelectionListener(this);
                Globals.mTree.addKeyListener(this);
                Globals.mMainForm.treeScrollPane.setViewportView(Globals.mTree);
                TreePath path = new TreePath(Globals.mRootNode);
                ((Item) Globals.mRootNode.getUserObject()).onExpand();
                Globals.mTree.expandPath(path);
                Globals.mTree.setSelectionPath(path);

                // Enable all toolbar buttons and menu items
                Globals.mMainForm.menuSaveAs.setEnabled(true);
                Globals.mMainForm.menuExpandAll.setEnabled(true);
                Globals.mMainForm.menuCollapseAll.setEnabled(true);
                Globals.mMainForm.toolSaveAs.setEnabled(true);
                Globals.mMainForm.toolExpandAll.setEnabled(true);
                Globals.mMainForm.toolCollapseAll.setEnabled(true);
            }
        } catch (Exception e) {
            new ErrorDialog(e);
        } finally {
            // Set the cursor back to normal even if an exception occurs.
            Globals.mMainForm.setCursor(null);
        }
    }

    /**
     * Saves the document with the name and format provided in standard Save As
     * dialog.
     */
    private void onSaveAs() {
        String fileName = Dialogs.saveDocument();
        if ("".equals(fileName) || Globals.mDocument == null) {
            return;
        }
        Globals.mMainForm.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        try {
            Globals.mDocument.save(fileName);
        } catch (Exception e) {
            new ErrorDialog(e);
        } finally {
            // Set the cursor back to normal even if an exception occurs.
            Globals.mMainForm.setCursor(null);
        }
    }

    /**
     * Expand all child nodes under the selected node.
     */
    private void onExpandAll() {
        TreePath path = Globals.mTree.getSelectionPath();
        if (path != null) {
            Globals.mMainForm.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
            expandAll(Globals.mTree, path, true);
            Globals.mMainForm.setCursor(null);
        }
    }

    /**
     * Collapse all child nodes under the selected node
     */
    private void onCollapseAll() {
        TreePath path = Globals.mTree.getSelectionPath();
        if (path != null) {
            Globals.mMainForm.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
            expandAll(Globals.mTree, path, false);
            Globals.mMainForm.setCursor(null);
        }
    }

    private void expandAll(JTree tree, TreePath parent, boolean expand) {
        // Traverse children.
        TreeNode node = (TreeNode) parent.getLastPathComponent();

        // Expansion or collapse must be done from the bottom-up
        if (expand) {
            tree.expandPath(parent);
        }

        if (node.getChildCount() >= 0) {
            for (Enumeration e = node.children(); e.hasMoreElements();) {
                TreeNode n = (TreeNode) e.nextElement();
                TreePath path = parent.pathByAddingChild(n);
                expandAll(tree, path, expand);
            }
        }

        if (!expand) {
            tree.collapsePath(parent);
        }
    }

    /**
     * Informs Item class, which provides GUI representation of a document node,
     * that the corresponding TreeNode is about being expanded.
     */
    public void treeWillExpand(TreeExpansionEvent event) throws ExpandVetoException {
        DefaultMutableTreeNode node = (DefaultMutableTreeNode) event.getPath().getLastPathComponent();
        if (node != null) {
            try {
                ((Item) node.getUserObject()).onExpand();
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
    }

    public void treeWillCollapse(TreeExpansionEvent event) throws ExpandVetoException {
    }

    /**
     * Informs Item class, which provides GUI representation of a document node,
     * that the corresponding TreeNode was selected.
     */
    public void valueChanged(TreeSelectionEvent e) {
        DefaultMutableTreeNode node = (DefaultMutableTreeNode) Globals.mTree.getLastSelectedPathComponent();

        if (node == null) {
            return;
        }
        try {
            // This operation can take some time so we set the Cursor to WaitCursor.
            Globals.mMainForm.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
            // Show the text contained by selected document node.
            Item selectedItem = (Item) node.getUserObject();
            Globals.mMainForm.textArea.setText(selectedItem.getText());
            Globals.mMainForm.textArea.moveCaretPosition(0);

            Globals.mMainForm.toolRemove.setEnabled(selectedItem.isRemovable());
            Globals.mMainForm.menuRemoveNode.setEnabled(selectedItem.isRemovable());

            // Restore cursor.
            Globals.mMainForm.setCursor(null);
        } catch (Exception ex) {
            Globals.mMainForm.textArea.setText("");
        }
    }

    /**
     * Removes the currently selected node.
     */
    private void onRemove() {
        DefaultMutableTreeNode node = (DefaultMutableTreeNode) Globals.mTree.getSelectionPath().getLastPathComponent();
        if (node != null) {
            try {
                ((Item) node.getUserObject()).remove();
            } catch (Exception e) {
            }
        }
    }

    /**
     * Show the About dialog
     */
    private void onAbout() {
        new About();
    }

    public void keyTyped(KeyEvent e) {
        if (e.getID() == KeyEvent.KEY_TYPED && e.getKeyChar() == 127) {
            onRemove();
        }
    }

    public void keyPressed(KeyEvent e) {
    }

    public void keyReleased(KeyEvent e) {
    }

    /**
     * Change the icon for the current node according to the node type.
     */
    private class OurCellRenderer extends DefaultTreeCellRenderer {

        public Component getTreeCellRendererComponent(
                JTree tree,
                Object value,
                boolean sel,
                boolean expanded,
                boolean leaf,
                int row,
                boolean hasFocus) {
            super.getTreeCellRendererComponent(
                    tree, value, sel,
                    expanded, leaf, row,
                    hasFocus);

            DefaultMutableTreeNode node = (DefaultMutableTreeNode) value;
            Object userObject = node.getUserObject();
            if (userObject instanceof Item) {
                ImageIcon icon = null;
                try {
                    icon = ((Item) userObject).getIcon();
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
                if (icon != null) {
                    setIcon(icon);
                }
            }
            return this;
        }
    }
}