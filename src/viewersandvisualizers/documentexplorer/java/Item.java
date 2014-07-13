/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package viewersandvisualizers.documentexplorer.java;

import com.aspose.words.*;

import javax.swing.*;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.TreeNode;
import javax.swing.tree.TreePath;
import java.lang.reflect.*;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

/**
* Base class used to provide GUI representation for document nodes.
*/
public class Item
{
    private Node mNode;
    private DefaultMutableTreeNode mTreeNode;
    private ImageIcon mIcon;

    private static ArrayList<java.lang.reflect.Field> mControlCharFields;
    private static Map<Integer, String> mNodeTypes;
    private static Map<Integer, String> mHeaderFooterTypes;
    private static Map<String, String> mItemSet;
    private static ArrayList mIconNames = new ArrayList();

    /**
    * Creates Item for the supplied document node.
    */
    public Item(Node node)
    {
        mNode = node;
    }

    /**
    * Returns the node in the document that this Item represents.
    */
    public Node getNode()
    {
        return mNode;
    }

    /**
    *  The display name for this Item. Can be customized by overriding this method in inheriting classes.
    */
    public String getName() throws Exception
	{
		return getNodeTypeString(mNode);
	}

    /**
    * The text of the corresponding document node.
    */
    public String getText() throws Exception
    {
        String text = mNode.getText();

        // Most control characters are converted to human readable form.
        // E.g. [!PageBreak!], [!Cell!], etc.
        for (Field fieldInfo : mControlCharFields)
        {
            if (fieldInfo.getType() == char.class && Modifier.isStatic(fieldInfo.getModifiers()))
            {
                Character ch = fieldInfo.getChar(null);

                // Represent a paragraph break using the special formatting marker. This makes the text easier to read.
                if(fieldInfo.getName().equals("PARAGRAPH_BREAK_CHAR"))
                   text = text.replace(ch.toString(), "�" + "\n"); // JTextArea lines are separated using simple "\n" character and not using system independent new line character.
                else
                   text = text.replace(ch.toString(), java.text.MessageFormat.format("[!{0}!]", fieldInfo.getName().replace("_CHAR", "")));
            }
        }

        // All break chars should be supplemented with line feeds
        text = text.replace("BREAK!]", "BREAK!]\n");
        return text;
    }

    /**
    * Creates a TreeNode for this item to be displayed in the Document Explorer TreeView control.
    */
    public DefaultMutableTreeNode getTreeNode() throws Exception
	{
		if (mTreeNode == null)
		{
			mTreeNode = new DefaultMutableTreeNode(this);
			if (!mIconNames.contains(getIconName()))
			{
					mIconNames.add(getIconName());
			}

			if (mNode instanceof CompositeNode && ((CompositeNode)mNode).getChildNodes().getCount() > 0)
			{
				mTreeNode.add(new DefaultMutableTreeNode("#dummy"));
			}
		}
		return mTreeNode;
	}

    /**
    * Returns the icon to display in the Document Explorer TreeView control.
    */
    public ImageIcon getIcon() throws Exception
	{
		if (mIcon == null)
		{
			mIcon = loadIcon(getIconName());
			if (mIcon == null)
				mIcon = loadIcon("Node");
		}
		return mIcon;
	}

    /**
    * The icon for this node can be customized by overriding this property in the inheriting classes.
    * The name represents name of .ico file without extension located in the Icons folder of the project.
    */
    protected String getIconName() throws Exception
	{
		return getClass().getSimpleName().replace("Item", "");
	}

    /**
    * Provides lazy on-expand loading of underlying tree nodes.
    */
    public void onExpand() throws Exception
	{
		if ("#dummy".equals(getTreeNode().getFirstChild().toString()))
		{
			getTreeNode().removeAllChildren();
			Globals.mTreeModel.reload(getTreeNode());
			for (Object o : ((CompositeNode)mNode).getChildNodes())
			{
				Node n = (Node)o;
				getTreeNode().add(Item.createItem(n).getTreeNode());
			}
		}
	}

    /**
    * Loads and returns an icon from the assembly resource stream.
    */
    private ImageIcon loadIcon(String iconName)
    {
        java.net.URL imgURL = MainForm.class.getResource("images/" + iconName + ".gif");
        if(imgURL != null)
                return new ImageIcon(imgURL);
        else
                return null;
    }

    /**
    * Removes this node from the document and the tree.
    */
	public void remove() throws Exception
	{
		if (this.isRemovable())
		{
            mNode.remove();
            TreeNode parent = mTreeNode.getParent();
            mTreeNode.removeFromParent();
            Globals.mTreeModel.reload(parent);
            TreePath path = new TreePath(Globals.mRootNode);
            Globals.mTree.setSelectionPath(path);
        }
	}

    /**
     * Returns if this node can be removed from the document. Some nodes such as the last paragraph in the
    * document cannot be removed.
    */
    public boolean isRemovable()
	{
		return true;
    }

    /**
    * Static ctor.
    */
    static
    {
		// Populate a list of node types along with their class implementation.
		mItemSet = new HashMap<String, String>();
		for(Class itemClass : DocumentItems.class.getDeclaredClasses())
		{
			try
			{
				String nodeTypeString = (String) itemClass.getField("NODE_TYPE_STRING").get(null);
				mItemSet.put(nodeTypeString, itemClass.getName());
			}
			catch (Exception e)
			{
				// IllegalAccessException, NoSuchFieldException or NoSuchMethodException - skip such exceptions if there are any.
			}
		}

		// Fill a list containing the information of each control char.
		mControlCharFields = new ArrayList<java.lang.reflect.Field>();
		Field[] fields = ControlChar.class.getFields();
		for(Field fieldInfo : fields)
		{
			if(fieldInfo.getType() == char.class && Modifier.isStatic(fieldInfo.getModifiers()))
			{
				if(!fieldInfo.getName().equals("SPACE_CHAR"))
                    mControlCharFields.add(fieldInfo);
			}
		}

		// Map node type integer values to their equivalent string name.
		mNodeTypes = new HashMap<Integer, String>();
		Field[] nodeTypefields = NodeType.class.getFields();

		for(Field fieldInfo : nodeTypefields)
		{
			if (fieldInfo.getType() == int.class && Modifier.isStatic(fieldInfo.getModifiers()))
			{
				try
				{
					int integerValue = fieldInfo.getInt(null);
					mNodeTypes.put(integerValue, fieldInfo.getName());
				}
				catch (IllegalAccessException e)
				{
					// Skip any invalid fields.
				}
			}
		}

		// Maps header/footer type integer values to string names.
		mHeaderFooterTypes = new HashMap<Integer, String>();
		fields = HeaderFooterType.class.getFields();

		for(Field fieldInfo : fields)
		{
			if(fieldInfo.getType() == int.class && Modifier.isStatic(fieldInfo.getModifiers()))
			{
				try
				{
					int integerValue = fieldInfo.getInt(null);
					mHeaderFooterTypes.put(integerValue, fieldInfo.getName());
				}
				catch (IllegalAccessException e)
				{
					// Skip any invalid fields.
				}
			}
		}
    }

    /**
    * Item class factory implementation.
    */
    public static Item createItem(Node node) throws ClassNotFoundException, NoSuchMethodException,
													IllegalAccessException, InvocationTargetException,
													InstantiationException
	{
		String typeName = getNodeTypeString(node);
		if (mItemSet.containsKey(typeName))
			return (Item)Class.forName(mItemSet.get(typeName)).
					getConstructor(DocumentItems.class, Node.class).
					newInstance(null, node);
		else
			return new Item(node);
	}

    /**
    * Object.toString method used by Tree.
    */
    public String toString()
    {
		// Introduced non-checked RuntimeException on purpose to not change Object.toString() signature
		try
		{
			return getName();
		}
		catch (Exception e)
		{
			throw new RuntimeException(e);
		}
	}

    /**
    * Convert numerical representation of the node type to string.
    */
    private static String getNodeTypeString(Node node)
    {
        int nodeType = node.getNodeType();
        if(mNodeTypes.containsKey(nodeType))
            return mNodeTypes.get(nodeType);
        else
            return "";
    }

    /**
    * Convert numerical representation of HeaderFooter integer type to string.
    */
    protected static String getHeaderFooterTypeAsString(HeaderFooter headerFooter) throws Exception
    {
		int headerFooterType = headerFooter.getHeaderFooterType();
		if(mHeaderFooterTypes.containsKey(headerFooterType))
			return mHeaderFooterTypes.get(headerFooterType);
		else
			return "";
    }
}