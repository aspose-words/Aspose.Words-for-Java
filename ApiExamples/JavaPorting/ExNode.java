// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import com.aspose.words.Node;
import org.testng.Assert;
import com.aspose.words.CompositeNode;
import com.aspose.ms.System.msString;
import com.aspose.ms.System.msConsole;
import com.aspose.words.NodeType;
import com.aspose.words.NodeCollection;
import com.aspose.words.Inline;
import com.aspose.words.TableCollection;
import com.aspose.words.Table;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.BreakType;
import com.aspose.words.Section;
import com.aspose.words.NodeList;
import java.util.Iterator;
import com.aspose.words.Body;
import com.aspose.words.SaveFormat;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.BitmapPal;
import java.awt.image.BufferedImage;
import com.aspose.words.INodeChangingCallback;
import com.aspose.words.NodeChangingArgs;
import com.aspose.words.NodeChangingAction;
import com.aspose.words.RunCollection;
import com.aspose.words.Shape;


@Test
public class ExNode extends ApiExampleBase
{
    @Test
    public void cloneCompositeNode() throws Exception
    {
        //ExStart
        //ExFor:Node
        //ExFor:Node.Clone
        //ExSummary:Shows how to clone composite nodes with and without their child nodes.
        Document doc = new Document();
        Paragraph para = doc.getFirstSection().getBody().getFirstParagraph();
        para.appendChild(new Run(doc, "Hello world!"));

        // Clone the paragraph and the child nodes
        Node cloneWithChildren = para.deepClone(true);

        Assert.assertTrue(((CompositeNode)cloneWithChildren).hasChildNodes());
        Assert.assertEquals("Hello world!", msString.trim(cloneWithChildren.getText()));

        // Clone the paragraph without its clild nodes
        Node cloneWithoutChildren = para.deepClone(false);

        Assert.assertFalse(((CompositeNode)cloneWithoutChildren).hasChildNodes());
        Assert.assertEquals("", msString.trim(cloneWithoutChildren.getText()));
        //ExEnd
    }

    @Test
    public void getParentNode() throws Exception
    {
        //ExStart
        //ExFor:Node.ParentNode
        //ExSummary:Shows how to access the parent node.
        Document doc = new Document();

        // Get the document's first paragraph and append a child node to it in the form of a run with text
        Paragraph para = doc.getFirstSection().getBody().getFirstParagraph();

        // When inserting a new node, the document that the node will belong to must be provided as an argument
        Run run = new Run(doc, "Hello world!");
        para.appendChild(run);

        // The node lineage can be traced back to the document itself
        Assert.assertEquals(para, run.getParentNode());
        Assert.assertEquals(doc.getFirstSection().getBody(), para.getParentNode());
        Assert.assertEquals(doc.getFirstSection(), doc.getFirstSection().getBody().getParentNode());
        Assert.assertEquals(doc, doc.getFirstSection().getParentNode());
        //ExEnd
    }

    @Test
    public void ownerDocument() throws Exception
    {
        //ExStart
        //ExFor:Node.Document
        //ExFor:Node.ParentNode
        //ExSummary:Shows how to create a node and set its owning document.
        // Open a file from disk
        Document doc = new Document();

        // Creating a new node of any type requires a document passed into the constructor
        Paragraph para = new Paragraph(doc);

        // The new paragraph node does not yet have a parent
        System.out.println("Paragraph has no parent node: " + (para.getParentNode() == null));

        // But the paragraph node knows its document
        System.out.println("Both nodes' documents are the same: " + (para.getDocument() == doc));

        // The fact that a node always belongs to a document allows us to access and modify 
        // properties that reference the document-wide data such as styles or lists
        para.getParagraphFormat().setStyleName("Heading 1");

        // Now add the paragraph to the main text of the first section
        doc.getFirstSection().getBody().appendChild(para);

        // The paragraph node is now a child of the Body node
        System.out.println("Paragraph has a parent node: " + (para.getParentNode() != null));
        //ExEnd

        Assert.assertEquals(doc, para.getDocument());
        Assert.assertNotNull(para.getParentNode());
    }

    @Test
    public void enumerateChildNodes() throws Exception
    {
        //ExStart
        //ExFor:Node
        //ExFor:NodeType
        //ExFor:CompositeNode
        //ExFor:CompositeNode.GetChild
        //ExFor:CompositeNode.ChildNodes
        //ExFor:CompositeNode.GetEnumerator
        //ExSummary:Shows how to enumerate immediate children of a CompositeNode using the enumerator provided by the ChildNodes collection.
        Document doc = new Document();

        Paragraph paragraph = (Paragraph)doc.getChild(NodeType.PARAGRAPH, 0, true);
        paragraph.appendChild(new Run(doc, "Hello world!"));
        paragraph.appendChild(new Run(doc, " Hello again!"));

        NodeCollection children = paragraph.getChildNodes();

        // Paragraph may contain children of various types such as runs, shapes and so on
        for (Node child : (Iterable<Node>) children)
            if (((child.getNodeType()) == (NodeType.RUN)))
            {
                Run run = (Run)child;
                System.out.println(run.getText());
            }
        //ExEnd

        Assert.assertEquals(NodeType.RUN, paragraph.getChild(NodeType.RUN, 0, true).getNodeType());
        Assert.assertEquals(2, paragraph.getChildNodes().getCount());
        Assert.assertEquals("Hello world! Hello again!", msString.trim(doc.getText()));
    }

    @Test
    public void indexChildNodes() throws Exception
    {
        //ExStart
        //ExFor:NodeCollection.Count
        //ExFor:NodeCollection.Item
        //ExSummary:Shows how to enumerate immediate children of a CompositeNode using indexed access.
        Document doc = new Document();
        Paragraph paragraph = (Paragraph)doc.getChild(NodeType.PARAGRAPH, 0, true);
        paragraph.appendChild(new Run(doc, "Hello world!"));

        NodeCollection children = paragraph.getChildNodes();

        for (int i = 0; i < children.getCount(); i++)
        {
            Node child = children.get(i);

            // Paragraph may contain children of various types such as runs, shapes and so on
            if (((child.getNodeType()) == (NodeType.RUN)))
            {
                Run run = (Run)child;
                System.out.println(run.getText());
            }
        }
        //ExEnd

        Assert.assertEquals(1, paragraph.getChildNodes().getCount());
    }

    //ExStart
    //ExFor:Node.NextSibling
    //ExFor:CompositeNode.FirstChild
    //ExFor:Node.IsComposite
    //ExFor:CompositeNode.IsComposite
    //ExFor:Node.NodeTypeToString
    //ExFor:Paragraph.NodeType
    //ExFor:Table.NodeType
    //ExFor:Node.NodeType
    //ExFor:Footnote.NodeType
    //ExFor:FormField.NodeType
    //ExFor:SmartTag.NodeType
    //ExFor:Cell.NodeType
    //ExFor:Row.NodeType
    //ExFor:Document.NodeType
    //ExFor:Comment.NodeType
    //ExFor:Run.NodeType
    //ExFor:Section.NodeType
    //ExFor:SpecialChar.NodeType
    //ExFor:Shape.NodeType
    //ExFor:FieldEnd.NodeType
    //ExFor:FieldSeparator.NodeType
    //ExFor:FieldStart.NodeType
    //ExFor:BookmarkStart.NodeType
    //ExFor:CommentRangeEnd.NodeType
    //ExFor:BuildingBlock.NodeType
    //ExFor:GlossaryDocument.NodeType
    //ExFor:BookmarkEnd.NodeType
    //ExFor:GroupShape.NodeType
    //ExFor:CommentRangeStart.NodeType
    //ExSummary:Shows how to efficiently visit all direct and indirect children of a composite node.
    @Test //ExSkip
    public void recurseAllNodes() throws Exception
    {
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // Any node that can contain child nodes, such as the document itself, is composite
        Assert.assertTrue(doc.isComposite());

        // Invoke the recursive function that will go through and print all the child nodes of a composite node
        traverseAllNodes(doc, 0);
    }

    /// <summary>
    /// Recursively traverses a node tree while printing the type of each node with an indent depending on depth as well as the contents of all inline nodes.
    /// </summary>
    @Test (enabled = false)
    public void traverseAllNodes(CompositeNode parentNode, int depth)
    {
        // Loop through immediate children of a node
        for (Node childNode = parentNode.getFirstChild(); childNode != null; childNode = childNode.getNextSibling())
        {
            msConsole.write($"{new string('\t', depth)}{Node.NodeTypeToString(childNode.NodeType)}");

            // Recurse into the node if it is a composite node
            if (childNode.isComposite())
            {
                msConsole.writeLine();
                traverseAllNodes((CompositeNode)childNode, depth + 1);
            }
            else if (childNode instanceof Inline)
            {
                System.out.println(" - \"{childNode.GetText().Trim()}\"");
            }
            else
            {
                msConsole.writeLine();
            }
        }
    }
    //ExEnd

    @Test
    public void removeNodes() throws Exception
    {

        //ExStart
        //ExFor:Node
        //ExFor:Node.NodeType
        //ExFor:Node.Remove
        //ExSummary:Shows how to remove all nodes of a specific type from a composite node.
        Document doc = new Document(getMyDir() + "Tables.docx");

        Assert.assertEquals(2, doc.getChildNodes(NodeType.TABLE, true).getCount());

        // Select the first child node in the body
        Node curNode = doc.getFirstSection().getBody().getFirstChild();

        while (curNode != null)
        {
            // Save the next sibling node as a variable in case we want to move to it after deleting this node
            Node nextNode = curNode.getNextSibling();

            // A section body can contain Paragraph and Table nodes
            // If the node is a Table, remove it from the parent
            if (((curNode.getNodeType()) == (NodeType.TABLE)))
                curNode.remove();

            // Continue going through child nodes until null (no more siblings) is reached
            curNode = nextNode;
        }

        Assert.assertEquals(0, doc.getChildNodes(NodeType.TABLE, true).getCount());
        //ExEnd
    }

    @Test
    public void enumNextSibling() throws Exception
    {

        //ExStart
        //ExFor:CompositeNode.FirstChild
        //ExFor:Node.NextSibling
        //ExFor:Node.NodeTypeToString
        //ExFor:Node.NodeType
        //ExSummary:Shows how to enumerate immediate child nodes of a composite node using NextSibling.
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // Loop starting from the first child until we reach null
        for (Node node = doc.getFirstSection().getBody().getFirstChild(); node != null; node = node.getNextSibling())
        {
            // Output the types of the nodes that we come across
            System.out.println(Node.nodeTypeToString(node.getNodeType()));
        }
        //ExEnd
    }

    @Test
    public void typedAccess() throws Exception
    {

        //ExStart
        //ExFor:Story.Tables
        //ExFor:Table.FirstRow
        //ExFor:Table.LastRow
        //ExFor:TableCollection
        //ExSummary:Shows how to use typed properties to access nodes of the document tree.
        Document doc = new Document(getMyDir() + "Tables.docx");

        // Quick typed access to all Table child nodes contained in the Body
        TableCollection tables = doc.getFirstSection().getBody().getTables();

        Assert.assertEquals(5, tables.get(0).getRows().getCount());
        Assert.assertEquals(4, tables.get(1).getRows().getCount());

        for (Table table : tables.<Table>OfType() !!Autoporter error: Undefined expression type )
        {
            // Quick typed access to the first row of the table
            table.getFirstRow()?.Remove();

            // Quick typed access to the last row of the table
            table.getLastRow()?.Remove();
        }

        // Each table has shrunk by two rows
        Assert.assertEquals(3, tables.get(0).getRows().getCount());
        Assert.assertEquals(2, tables.get(1).getRows().getCount());
        //ExEnd
    }

    @Test
    public void removeChild() throws Exception
    {
        //ExStart
        //ExFor:CompositeNode.LastChild
        //ExFor:Node.PreviousSibling
        //ExFor:CompositeNode.RemoveChild
        //ExSummary:Shows how to use of methods of Node and CompositeNode to remove a section before the last section in the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a second section by inserting a section break and add text to both sections
        builder.writeln("Section 1 text.");
        builder.insertBreak(BreakType.SECTION_BREAK_CONTINUOUS);
        builder.writeln("Section 2 text.");

        // Both sections are siblings of each other
        Section lastSection = (Section)doc.getLastChild();
        Section firstSection = (Section)lastSection.getPreviousSibling();

        // Remove a section based on its sibling relationship with another section
        if (lastSection.getPreviousSibling() != null)
            doc.removeChild(firstSection);

        // The section we removed was the first one, leaving the document with only the second
        Assert.assertEquals("Section 2 text.", msString.trim(doc.getText()));
        //ExEnd
    }

    @Test
    public void selectCompositeNodes() throws Exception
    {
        //ExStart
        //ExFor:CompositeNode.SelectSingleNode
        //ExFor:CompositeNode.SelectNodes
        //ExFor:NodeList.GetEnumerator
        //ExFor:NodeList.ToArray
        //ExSummary:Shows how to select certain nodes by using an XPath expression.
        Document doc = new Document(getMyDir() + "Tables.docx");

        // This expression will extract all paragraph nodes which are descendants of any table node in the document
        // This will return any paragraphs which are in a table
        NodeList nodeList = doc.selectNodes("//Table//Paragraph");

        // Iterate through the list with an enumerator and print the contents of every paragraph in each cell of the table
        int index = 0;

        Iterator<Node> e = nodeList.iterator();
        try /*JAVA: was using*/
    	{
            while (e.hasNext())
                System.out.println("Table paragraph index {index++}, contents: \"{e.Current.GetText().Trim()}\"");
    	}
        finally { if (e != null) e.close(); }

        // This expression will select any paragraphs that are direct children of any body node in the document
        nodeList = doc.selectNodes("//Body/Paragraph");

        // We can treat the list as an array too
        Assert.assertEquals(4, nodeList.toArray().length);

        // Use SelectSingleNode to select the first result of the same expression as above
        Node node = doc.selectSingleNode("//Body/Paragraph");

        Assert.assertEquals(Paragraph.class, node.getClass());
        //ExEnd
    }

    @Test
    public void testNodeIsInsideField() throws Exception
    {
        //ExStart:
        //ExFor:CompositeNode.SelectNodes
        //ExSummary:Shows how to test if a node is inside a field by using an XPath expression.
        Document doc = new Document(getMyDir() + "Mail merge destination - Northwind employees.docx");

        // Evaluate the XPath expression. The resulting NodeList will contain all nodes found inside a field a field (between FieldStart 
        // and FieldEnd exclusive). There can however be FieldStart and FieldEnd nodes in the list if there are nested fields 
        // in the path. Currently does not find rare fields in which the FieldCode or FieldResult spans across multiple paragraphs
        NodeList resultList =
            doc.selectNodes("//FieldStart/following-sibling::node()[following-sibling::FieldEnd]");

        // Check if the specified run is one of the nodes that are inside the field
        System.out.println("Contents of the first Run node that's part of a field: {resultList.First(n => n.NodeType == NodeType.Run).GetText().Trim()}");
        //ExEnd
    }

    @Test
    public void createAndAddParagraphNode() throws Exception
    {
        Document doc = new Document();

        Paragraph para = new Paragraph(doc);

        Section section = doc.getLastSection();
        section.getBody().appendChild(para);
    }

    @Test
    public void removeSmartTagsFromCompositeNode() throws Exception
    {
        //ExStart
        //ExFor:CompositeNode.RemoveSmartTags
        //ExSummary:Removes all smart tags from descendant nodes of the composite node.
        Document doc = new Document(getMyDir() + "Smart tags.doc");
        Assert.assertEquals(8, doc.getChildNodes(NodeType.SMART_TAG, true).getCount());

        // Remove smart tags from the whole document
        doc.removeSmartTags();

        Assert.assertEquals(0, doc.getChildNodes(NodeType.SMART_TAG, true).getCount());
        //ExEnd
    }

    @Test
    public void getIndexOfNode() throws Exception
    {
        //ExStart
        //ExFor:CompositeNode.IndexOf
        //ExSummary:Shows how to get the index of a given child node from its parent.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Get the body of the first section in the document
        Body body = doc.getFirstSection().getBody();

        // Retrieve the index of the last paragraph in the body
        Assert.assertEquals(24, body.getChildNodes().indexOf(body.getLastParagraph()));
        //ExEnd
    }

    @Test
    public void convertNodeToHtmlWithDefaultOptions() throws Exception
    {
        //ExStart
        //ExFor:Node.ToString(SaveFormat)
        //ExFor:Node.ToString(SaveOptions)
        //ExSummary:Exports the content of a node to String in HTML format.
        Document doc = new Document(getMyDir() + "Document.docx");

        // Extract the last paragraph in the document to convert to HTML
        Node node = doc.getLastSection().getBody().getLastParagraph();

        // When ToString is called using the html SaveFormat overload then the node is converted directly to html
        Assert.assertEquals("<p style=\"margin-top:0pt; margin-bottom:8pt; line-height:108%; font-size:12pt\">" +
                        "<span style=\"font-family:'Times New Roman'\">Hello World!</span>" +
                        "</p>", node.toString(SaveFormat.HTML));

        // We can also modify the result of this conversion using a SaveOptions object
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportRelativeFontSize(true);

        Assert.assertEquals("<p style=\"margin-top:0pt; margin-bottom:8pt; line-height:108%\">" +
                        "<span style=\"font-family:'Times New Roman'\">Hello World!</span>" +
                        "</p>", node.toString(saveOptions));
        //ExEnd
    }

    @Test
    public void typedNodeCollectionToArray() throws Exception
    {
        //ExStart
        //ExFor:ParagraphCollection.ToArray
        //ExSummary:Shows how to create an array from a NodeCollection.
        // You can use ToArray to return a typed array of nodes
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        Paragraph[] paras = doc.getFirstSection().getBody().getParagraphs().toArray();

        Assert.assertEquals(22, paras.length);
        //ExEnd
    }

    @Test
    public void nodeEnumerationHotRemove() throws Exception
    {
        //ExStart
        //ExFor:ParagraphCollection.ToArray
        //ExSummary:Shows how to use "hot remove" to remove a node during enumeration.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.writeln("The first paragraph");
        builder.writeln("The second paragraph");
        builder.writeln("The third paragraph");
        builder.writeln("The fourth paragraph");

        // Hot remove allows a node to be removed from a live collection and have the enumeration continue
        for (Paragraph para : doc.getFirstSection().getBody().getParagraphs().toArray())
            if (para.getRange().getText().contains("third"))
                para.remove();
        
        Assert.assertFalse(doc.getText().contains("The third paragraph"));
        //ExEnd
    }

    //ExStart
    //ExFor:CompositeNode.CreateNavigator
    //ExSummary:Shows how to create an XPathNavigator and use it to traverse and read nodes.
    @Test //ExSkip
    public void nodeXPathNavigator() throws Exception
    {
        Document doc = new Document();

        // A document is a composite node so we can make a navigator straight away
        XPathNavigator navigator = doc.CreateNavigator();

        // Our root is the document node with 1 child, which is the first section
        if (navigator != null)
        {
            Assert.assertEquals("Document", navigator.Name);
            Assert.assertEquals(false, navigator.MoveToNext());
            Assert.assertEquals(1, navigator.SelectChildren(XPathNodeType.All).Count);

            // The document tree has the document, first section, body and first paragraph as nodes, with each being an only child of the previous
            // We can add a few more to give the tree some branches for the navigator to traverse
            DocumentBuilder docBuilder = new DocumentBuilder(doc);
            docBuilder.write("Section 1, Paragraph 1. ");
            docBuilder.insertParagraph();
            docBuilder.write("Section 1, Paragraph 2. ");
            doc.appendChild(new Section(doc));
            docBuilder.moveToSection(1);
            docBuilder.write("Section 2, Paragraph 1. ");

            // Use our navigator to print a map of all the nodes in the document to the console
            StringBuilder stringBuilder = new StringBuilder();
            mapDocument(navigator, stringBuilder, 0);
            msConsole.write(stringBuilder.toString());
            testNodeXPathNavigator(stringBuilder.toString(), doc); //ExSkip
        }
    }

    /// <summary>
    /// This will traverse all children of a composite node and map the structure in the style of a directory tree.
    /// Amount of space indentation indicates depth relative to initial node. Only runs will have their values printed.
    /// </summary>
    private static void mapDocument(XPathNavigator navigator, StringBuilder stringBuilder, int depth)
    {
        do
        {
            msStringBuilder.append(stringBuilder, ' ', depth);
            msStringBuilder.append(stringBuilder, navigator.Name + ": ");

            if ("Run".equals(navigator.Name))
            {
                msStringBuilder.append(stringBuilder, navigator.Value);
            }

            stringBuilder.append('\n');

            if (navigator.HasChildren)
            {
                navigator.MoveToFirstChild();
                mapDocument(navigator, stringBuilder, depth + 1);
                navigator.MoveToParent();
            }
        } while (navigator.MoveToNext());
    }
    //ExEnd

    private void testNodeXPathNavigator(String navigatorResult, Document doc)
    {
        for (Run run : doc.getChildNodes(NodeType.RUN, true).toArray().<Run>OfType() !!Autoporter error: Undefined expression type )
            Assert.assertTrue(navigatorResult.contains(msString.trim(run.getText())));
    }

    //ExStart
    //ExFor:NodeChangingAction
    //ExFor:NodeChangingArgs.Action
    //ExFor:NodeChangingArgs.NewParent
    //ExFor:NodeChangingArgs.OldParent
    //ExSummary:Shows how to use a NodeChangingCallback to monitor changes to the document tree as it is edited.
    @Test //ExSkip
    public void nodeChangingCallback() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the NodeChangingCallback attribute to a custom printer 
        doc.setNodeChangingCallback(new NodeChangingPrinter());

        // All node additions and removals will be printed to the console as we edit the document
        builder.writeln("Hello world!");
        builder.startTable();
        builder.insertCell();
        builder.write("Cell 1");
        builder.insertCell();
        builder.write("Cell 2");
        builder.endTable();

                builder.insertImage(BitmapPal.loadNativeImage(getImageDir() + "Logo.jpg"));
                    
        builder.getCurrentParagraph().getParentNode().removeAllChildren();
    }

    /// <summary>
    /// Prints every node insertion/removal as it takes place in the document.
    /// </summary>
    private static class NodeChangingPrinter implements INodeChangingCallback
    {
        public void /*INodeChangingCallback.*/nodeInserting(NodeChangingArgs args)
        {
            Assert.assertEquals(NodeChangingAction.INSERT, args.getAction());
            Assert.assertEquals(null, args.getOldParent());
        }

        public void /*INodeChangingCallback.*/nodeInserted(NodeChangingArgs args)
        {
            Assert.assertEquals(NodeChangingAction.INSERT, args.getAction());
            Assert.assertNotNull(args.getNewParent());

            System.out.println("Inserted node:");
            System.out.println("\tType:\t{args.Node.NodeType}");

            if (!"".equals(msString.trim(args.getNode().getText())))
            {
                System.out.println("\tText:\t\"{args.Node.GetText().Trim()}\"");
            }

            System.out.println("\tHash:\t{args.Node.GetHashCode()}");
            System.out.println("\tParent:\t{args.NewParent.NodeType} ({args.NewParent.GetHashCode()})");
        }

        public void /*INodeChangingCallback.*/nodeRemoving(NodeChangingArgs args)
        {
            Assert.assertEquals(NodeChangingAction.REMOVE, args.getAction());
        }

        public void /*INodeChangingCallback.*/nodeRemoved(NodeChangingArgs args)
        {
            Assert.assertEquals(NodeChangingAction.REMOVE, args.getAction());
            Assert.assertNull(args.getNewParent());

            System.out.println("Removed node: {args.Node.NodeType} ({args.Node.GetHashCode()})");
        }
    }
    //ExEnd

    @Test
    public void nodeCollection() throws Exception
    {
        //ExStart
        //ExFor:NodeCollection.Contains(Node)
        //ExFor:NodeCollection.Insert(Int32,Node)
        //ExFor:NodeCollection.Remove(Node)
        //ExSummary:Shows how to work with a NodeCollection.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The normal way to insert Runs into a document is to add text using a DocumentBuilder
        builder.write("Run 1. ");
        builder.write("Run 2. ");

        // Every .Write() invocation creates a new Run, which is added to the parent Paragraph's RunCollection
        RunCollection runs = doc.getFirstSection().getBody().getFirstParagraph().getRuns();
        Assert.assertEquals(2, runs.getCount());

        // We can insert a node into the RunCollection manually to achieve the same effect
        Run newRun = new Run(doc, "Run 3. ");
        runs.insert(3, newRun);

        Assert.assertTrue(runs.contains(newRun));
        Assert.assertEquals("Run 1. Run 2. Run 3.", msString.trim(doc.getText()));

        // Text can also be deleted from the document by accessing individual Runs via the RunCollection and editing or removing them
        Run run = runs.get(1);
        runs.remove(run);
        Assert.assertEquals("Run 1. Run 3.", msString.trim(doc.getText()));

        Assert.assertNotNull(run);
        Assert.assertFalse(runs.contains(run));
        //ExEnd
    }

    @Test
    public void nodeList() throws Exception
    {
        //ExStart
        //ExFor:NodeList.Count
        //ExFor:NodeList.Item(System.Int32)
        //ExSummary:Shows how to use XPaths to navigate a NodeList.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some nodes with a DocumentBuilder
        builder.writeln("Hello world!");

        builder.startTable();
        builder.insertCell();
        builder.write("Cell 1");
        builder.insertCell();
        builder.write("Cell 2");
        builder.endTable();

                builder.insertImage(BitmapPal.loadNativeImage(getImageDir() + "Logo.jpg"));
                    
        // Get all run nodes, of which we put 3 in the entire document
        NodeList nodeList = doc.selectNodes("//Run");
        Assert.assertEquals(3, nodeList.getCount());

        // Using a double forward slash, select all Run nodes that are indirect descendants of a Table node,
        // which would in this case be the runs inside the two cells we inserted
        nodeList = doc.selectNodes("//Table//Run");
        Assert.assertEquals(2, nodeList.getCount());

        // Single forward slashes specify direct descendant relationships,
        // of which we skipped quite a few by using double slashes
        Assert.assertEquals(doc.selectNodes("//Table//Run"), doc.selectNodes("//Table/Row/Cell/Paragraph/Run"));

        // We can access the actual nodes via a NodeList too
        nodeList = doc.selectNodes("//Shape");
        Assert.assertEquals(1, nodeList.getCount());
        Shape shape = (Shape)nodeList.get(0);
        Assert.assertTrue(shape.hasImage());
        //ExEnd
    }
}
