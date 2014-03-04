/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
import com.aspose.words.*;

import com.aspose.words.Document;
import com.aspose.words.License;

import javax.swing.*;
import javax.swing.event.TreeExpansionEvent;
import javax.swing.event.TreeSelectionEvent;
import javax.swing.event.TreeSelectionListener;
import javax.swing.event.TreeWillExpandListener;
import javax.swing.tree.*;
import java.awt.*;
import java.awt.event.*;
import java.io.File;
import java.util.Enumeration;

/**
* The main form of the DocumentExplorer demo.
*
* DocumentExplorer allows to open DOC, DOT, DOCX, XML, WML, RTF, ODT,
* OTT, HTML, XHTML and MHTML files using Aspose.Words.
*
* Once a document is opened, you can explore its object model in the tree.
* You can also save the document into DOC, DOCX, ODF, EPUB, PDF, SWF, RTF, WordML,
* HTML, MHTML and plain text formats.

*/
public class MainForm extends JDialog implements TreeWillExpandListener, TreeSelectionListener, KeyListener
{
	private JPanel contentPane;
	private JPanel jPanel2;
	private JTextPane textPane1;
	private JScrollPane treeScrollPane;
	private JMenuItem menuSaveAs;
	private JMenuItem menuExpandAll;
	private JMenuItem menuCollapseAll;
	private JMenuItem menuRemoveNode;
	private JButton toolSaveAs;
	private JButton toolExpandAll;
	private JButton toolCollapseAll;
	private JButton toolRemove;

	/**
	* Ctor.
	*/
	public MainForm() throws Exception
	{
		initComponents();
		setContentPane(contentPane);
		setModal(true);

		// Search for an Aspose.Words license in the application directory.
		// The File.Exists check is only needed in this demo so it will work
		// both when the license file is present as well as when it's missing.
		// In your real application you just need to call the SetLicense method.
		File licenseFile = new File(System.getProperty("user.dir") + "\\Aspose.Words.lic");
		if (licenseFile.exists())
		{
			// This shows how to license Aspose.Words.
			// If you don't specify a license, Aspose.Words works in evaluation mode.
			License license = new License();
			license.setLicense(licenseFile.getAbsolutePath());
		}
		Globals.mMainForm = this;
        Globals.mMainForm.setTitle(Globals.APPLICATION_TITLE);
        Frame parentFrame = (Frame)Globals.mMainForm.getOwner();
        parentFrame.setIconImage(Utils.createImageIcon("images/App.gif").getImage());
	}

	/**
	* Initialize Menu and Toolbar
	*/
	private void initComponents()
	{
		setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);
		addWindowListener(new WindowAdapter()
		{
			public void windowClosing(WindowEvent e)
			{
				onClose();
			}
		});

		// ToolBar
		jPanel2.setLayout(new BorderLayout());

		JToolBar jToolBar = new JToolBar();
		JButton jToolbarButton;
		jToolbarButton = new JButton(Utils.createImageIcon("images/tlb_1.gif"));
		jToolbarButton.setEnabled(true);
		jToolbarButton.setToolTipText("Open Document");
		jToolbarButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt) {onOpen();}
		});
		jToolBar.add(jToolbarButton);

		toolSaveAs = new JButton(Utils.createImageIcon("images/tlb_2.gif"));
		toolSaveAs.setEnabled(false);
		toolSaveAs.setToolTipText("Save Document As ...");
		toolSaveAs.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt) {onSaveAs();}
		});
		jToolBar.add(toolSaveAs);

		toolExpandAll = new JButton(Utils.createImageIcon("images/tlb_3.gif"));
		toolExpandAll.setEnabled(false);
		toolExpandAll.setToolTipText("Expand All");
		toolExpandAll.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt) {onExpandAll();}
		});
		jToolBar.add(toolExpandAll);

		toolCollapseAll = new JButton(Utils.createImageIcon("images/tlb_4.gif"));
		toolCollapseAll.setEnabled(false);
		toolCollapseAll.setToolTipText("Collapse All");
		toolCollapseAll.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt) {onCollapseAll();}
		});
		jToolBar.add(toolCollapseAll);

		toolRemove = new JButton(Utils.createImageIcon("images/tlb_5.gif"));
		toolRemove.setEnabled(false);
		toolRemove.setToolTipText("Remove Node");
		toolRemove.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt) {onRemove();}
		});
		jToolBar.add(toolRemove);

		jPanel2.add(jToolBar);

		//Menu
		JMenuBar jMainMenu;

		JMenu jMenu;
		JMenuItem jMenuItem;

		jMainMenu = new JMenuBar();

		// File Menu
		jMenu = new JMenu();
		jMenu.setMnemonic('F');
		jMenu.setText("File");
		jMenuItem = new JMenuItem();
		jMenuItem.setMnemonic('O');
		jMenuItem.setText("Open");
		jMenuItem.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt) {onOpen();}
		});
		jMenu.add(jMenuItem);

		menuSaveAs = new JMenuItem();
		menuSaveAs.setMnemonic('A');
		menuSaveAs.setText("Save As");
		menuSaveAs.setEnabled(false);
		menuSaveAs.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt) {onSaveAs();}
		});
		jMenu.add(menuSaveAs);

		jMenu.add(new JSeparator(JSeparator.HORIZONTAL));

		jMenuItem = new JMenuItem();
		jMenuItem.setMnemonic('X');
		jMenuItem.setText("Exit");
		jMenuItem.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt) {onClose();}
		});
		jMenu.add(jMenuItem);

		jMainMenu.add(jMenu);

		// Edit Menu
		jMenu = new JMenu();
		jMenu.setText("Edit");
		menuRemoveNode = new JMenuItem();
		menuRemoveNode.setText("Remove Node");
		menuRemoveNode.setEnabled(false);
		menuRemoveNode.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt) {onRemove();}
		});
		jMenu.add(menuRemoveNode);
		jMainMenu.add(jMenu);

		// View Menu
		jMenu = new JMenu();
		jMenu.setMnemonic('V');
		jMenu.setText("View");
		menuExpandAll = new JMenuItem();
		menuExpandAll.setMnemonic('E');
		menuExpandAll.setText("Expand All");
		menuExpandAll.setEnabled(false);
		menuExpandAll.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt) {onExpandAll();}
		});
		jMenu.add(menuExpandAll);

		menuCollapseAll = new JMenuItem();
		menuCollapseAll.setMnemonic('C');
		menuCollapseAll.setText("Collapse All");
		menuCollapseAll.setEnabled(false);
		menuCollapseAll.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt) {onCollapseAll();}
		});
		jMenu.add(menuCollapseAll);
		jMainMenu.add(jMenu);

		// Help Menu
		jMenu = new JMenu();
		jMenu.setMnemonic('H');
		jMenu.setText("Help");

		jMenuItem = new JMenuItem();
		jMenuItem.setMnemonic('A');
		jMenuItem.setText("About");
		jMenuItem.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt) {onAbout();}
		});
		jMenu.add(jMenuItem);

		jMainMenu.add(jMenu);

		setJMenuBar(jMainMenu);
	}

	private void onClose()
	{
		dispose();
	}

	/**
	* Opens a document with the name and format provided in a standard Save As dialog.
	*/
	private void onOpen()
	{
		try
		{
			String fileName = Dialogs.openDocument();
			if (!"".equals(fileName))
			{
				setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
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
				treeScrollPane.setViewportView(Globals.mTree);
				TreePath path = new TreePath(Globals.mRootNode);
				((Item) Globals.mRootNode.getUserObject()).onExpand();
				Globals.mTree.expandPath(path);
				Globals.mTree.setSelectionPath(path);

				// Enable all toolbar buttons and menu items
				menuSaveAs.setEnabled(true);
				menuExpandAll.setEnabled(true);
				menuCollapseAll.setEnabled(true);
				toolSaveAs.setEnabled(true);
				toolExpandAll.setEnabled(true);
				toolCollapseAll.setEnabled(true);
			}
		}
		catch (Exception e)
		{
			ExceptionDialog dialog = new ExceptionDialog(e);
			dialog.pack();
			dialog.setVisible(true);
		}

        finally
        {
            // Set the cursor back to normal even if an exception occurs.
           	setCursor(null);
        }
	}

	/**
	* Saves the document with the name and format provided in standard Save As dialog.
	*/
	private void onSaveAs()
	{
		String fileName = Dialogs.saveDocument();
		if ("".equals(fileName) || Globals.mDocument == null)
		{
			return;
		}
		setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

		try
		{
		    Globals.mDocument.save(fileName);
		}
		catch (Exception e)
		{
			ExceptionDialog dialog = new ExceptionDialog(e);
			dialog.pack();
			dialog.setVisible(true);
		}

        finally
        {
            // Set the cursor back to normal even if an exception occurs.
		    setCursor(null);
        }
	}

	/**
	* Expand all child nodes under the selected node.
	*/
	private void onExpandAll()
	{
		TreePath path = Globals.mTree.getSelectionPath();
		if (path != null)
		{
			setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
			expandAll(Globals.mTree, path, true);
			setCursor(null);
		}
	}

	/**
	* Collapse all child nodes under the selected node
	*/
	private void onCollapseAll()
	{
		TreePath path = Globals.mTree.getSelectionPath();
		if (path != null)
		{
			setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
			expandAll(Globals.mTree, path, false);
			setCursor(null);
		}
	}

	private void expandAll(JTree tree, TreePath parent, boolean expand)
	{
		// Traverse children.
		TreeNode node = (TreeNode) parent.getLastPathComponent();

		// Expansion or collapse must be done from the bottom-up
		if (expand)
		{
			tree.expandPath(parent);
		}

		if (node.getChildCount() >= 0)
		{
			for (Enumeration e = node.children(); e.hasMoreElements();)
			{
				TreeNode n = (TreeNode) e.nextElement();
				TreePath path = parent.pathByAddingChild(n);
				expandAll(tree, path, expand);
			}
		}

		if (!expand)
		{
			tree.collapsePath(parent);
		}
	}

	/**
	* Informs Item class, which provides GUI representation of a document node,
	* that the corresponding TreeNode is about being expanded.
	*/
	public void treeWillExpand(TreeExpansionEvent event) throws ExpandVetoException
	{
		DefaultMutableTreeNode node = (DefaultMutableTreeNode) event.getPath().getLastPathComponent();
		if (node != null)
		{
			try
			{
				((Item) node.getUserObject()).onExpand();
			}
			catch (Exception e)
			{
				throw new RuntimeException(e);
			}
		}
	}

	public void treeWillCollapse(TreeExpansionEvent event) throws ExpandVetoException
	{
	}

	/**
	* Informs Item class, which provides GUI representation of a document node,
	* that the corresponding TreeNode was selected.
	*/
	public void valueChanged(TreeSelectionEvent e)
	{
		DefaultMutableTreeNode node = (DefaultMutableTreeNode) Globals.mTree.getLastSelectedPathComponent();

		if (node == null) return;
		try
		{
			// This operation can take some time so we set the Cursor to WaitCursor.
			setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
			// Show the text contained by selected document node.
			Item selectedItem = (Item) node.getUserObject();
			textPane1.setText(selectedItem.getText());
			textPane1.moveCaretPosition(0);

			toolRemove.setEnabled(selectedItem.isRemovable());
			menuRemoveNode.setEnabled(selectedItem.isRemovable());

			// Restore cursor.
			setCursor(null);
		}
		catch (Exception ex)
		{
			textPane1.setText("");
		}
	}

	/**
	* Removes the currently selected node.
	*/
	private void onRemove()
	{
		DefaultMutableTreeNode node = (DefaultMutableTreeNode) Globals.mTree.getSelectionPath().getLastPathComponent();
		if (node != null)
		{
			try
			{
				((Item) node.getUserObject()).remove();
			}
			catch (Exception e)
			{
			}
		}
	}

	/**
	* Show the About dialog
	*/
	private void onAbout()
	{
		AboutForm dialog = new AboutForm();
		dialog.pack();

        // Get the screen size
		Toolkit toolkit = Toolkit.getDefaultToolkit();
		Dimension screenSize = toolkit.getScreenSize();

		// Calculate the frame location
		int x = (screenSize.width - dialog.getWidth()) / 2;
		int y = (screenSize.height - dialog.getHeight()) / 2;

		// Set the new frame location
		dialog.setLocation(x, y);
		dialog.setVisible(true);
	}

	public void keyTyped(KeyEvent e)
	{
		if (e.getID() == KeyEvent.KEY_TYPED && e.getKeyChar() == 127)
		{
			onRemove();
		}
	}

	public void keyPressed(KeyEvent e)
	{
	}

	public void keyReleased(KeyEvent e)
	{
	}

	/**
	* Change the icon for the current node according to the node type.
	*/
	private class OurCellRenderer extends DefaultTreeCellRenderer
	{
		public Component getTreeCellRendererComponent(
				JTree tree,
				Object value,
				boolean sel,
				boolean expanded,
				boolean leaf,
				int row,
				boolean hasFocus)
		{
			super.getTreeCellRendererComponent(
					tree, value, sel,
					expanded, leaf, row,
					hasFocus);

			DefaultMutableTreeNode node = (DefaultMutableTreeNode) value;
			Object userObject = node.getUserObject();
			if (userObject instanceof Item)
			{
				ImageIcon icon = null;
				try
				{
					icon = ((Item) userObject).getIcon();
				}
				catch (Exception e)
				{
					throw new RuntimeException(e);
				}
				if (icon != null)
				{
					setIcon(icon);
				}
			}
			return this;
		}
	}

	public static void main(String[] args) throws Exception
	{
		MainForm dialog = new MainForm();
		dialog.pack();

		// Get the screen size
		Toolkit toolkit = Toolkit.getDefaultToolkit();
		Dimension screenSize = toolkit.getScreenSize();

		// Calculate the frame location
		int x = (screenSize.width - dialog.getWidth()) / 2;
		int y = (screenSize.height - dialog.getHeight()) / 2;

		// Set the new frame location
		dialog.setLocation(x, y);
		dialog.setVisible(true);
		System.exit(0);
	}
}