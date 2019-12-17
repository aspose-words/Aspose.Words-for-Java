// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.ms.System.msConsole;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.NodeType;
import com.aspose.words.NodeCollection;
import com.aspose.words.Section;
import com.aspose.words.Body;
import com.aspose.words.TableCollection;
import com.aspose.words.Table;
import com.aspose.words.NodeList;
import java.util.Iterator;
import com.aspose.words.SaveFormat;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.ExportHeadersFootersMode;
import com.aspose.words.DocumentBuilder;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.BitmapPal;
import java.awt.image.BufferedImage;
import com.aspose.words.INodeChangingCallback;
import com.aspose.words.NodeChangingArgs;
import com.aspose.words.NodeChangingAction;
import com.aspose.ms.System.msString;
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
        // Create a new empty document.
        Document doc = new Document();

        // Add some text to the first paragraph
        Paragraph para = doc.getFirstSection().getBody().getFirstParagraph();
        para.appendChild(new Run(doc, "Some text"));

        // Clone the paragraph and the child nodes.
        Node cloneWithChildren = para.deepClone(true);
        // Only clone the paragraph and no child nodes.
        Node cloneWithoutChildren = para.deepClone(false);
        //ExEnd

        Assert.assertTrue(((CompositeNode) cloneWithChildren).hasChildNodes());
        Assert.assertFalse(((CompositeNode) cloneWithoutChildren).hasChildNodes());
    }

    @Test
    public void getParentNode() throws Exception
    {
        //ExStart
        //ExFor:Node.ParentNode
        //ExSummary:Shows how to access the parent node.
        // Create a new empty document. It has one section.
        Document doc = new Document();

        // The section is the first child node of the document.
        Node section = doc.getFirstChild();

        // The section's parent node is the document.
        msConsole.writeLine("Section parent is the document: " + (doc == section.getParentNode()));
        //ExEnd

        msAssert.areEqual(doc, section.getParentNode());
    }

    @Test
    public void ownerDocument() throws Exception
    {
        //ExStart
        //ExFor:Node.Document
        //ExFor:Node.ParentNode
        //ExSummary:Shows that when you create any node, it requires a document that will own the node.
        // Open a file from disk.
        Document doc = new Document();

        // Creating a new node of any type requires a document passed into the constructor.
        Paragraph para = new Paragraph(doc);

        // The new paragraph node does not yet have a parent.
        msConsole.writeLine("Paragraph has no parent node: " + (para.getParentNode() == null));

        // But the paragraph node knows its document.
        msConsole.writeLine("Both nodes' documents are the same: " + (para.getDocument() == doc));

        // The fact that a node always belongs to a document allows us to access and modify 
        // properties that reference the document-wide data such as styles or lists.
        para.getParagraphFormat().setStyleName("Heading 1");

        // Now add the paragraph to the main text of the first section.
        doc.getFirstSection().getBody().appendChild(para);

        // The paragraph node is now a child of the Body node.
        msConsole.writeLine("Paragraph has a parent node: " + (para.getParentNode() != null));
        //ExEnd

        msAssert.areEqual(doc, para.getDocument());
        Assert.assertNotNull(para.getParentNode());
    }

    @Test
    public void enumerateChildNodes() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:Node
        //ExFor:NodeType
        //ExFor:CompositeNode
        //ExFor:CompositeNode.GetChild
        //ExSummary:Shows how to extract a specific child node from a CompositeNode by using the GetChild method and passing the NodeType and index.
        Paragraph paragraph = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);
        //ExEnd

        //ExStart
        //ExFor:CompositeNode.ChildNodes
        //ExFor:CompositeNode.GetEnumerator
        //ExSummary:Shows how to enumerate immediate children of a CompositeNode using the enumerator provided by the ChildNodes collection.
        NodeCollection children = paragraph.getChildNodes();
        for (Node child : (Iterable<Node>) children)
        {
            // Paragraph may contain children of various types such as runs, shapes and so on.
            if (((child.getNodeType()) == (NodeType.RUN)))
            {
                // Say we found the node that we want, do something useful.
                Run run = (Run) child;
                msConsole.writeLine(run.getText());
            }
        }

        //ExEnd
    }

    @Test
    public void indexChildNodes() throws Exception
    {
        Document doc = new Document();
        Paragraph paragraph = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        //ExStart
        //ExFor:NodeCollection.Count
        //ExFor:NodeCollection.Item
        //ExSummary:Shows how to enumerate immediate children of a CompositeNode using indexed access.
        NodeCollection children = paragraph.getChildNodes();
        for (int i = 0; i < children.getCount(); i++)
        {
            Node child = children.get(i);

            // Paragraph may contain children of various types such as runs, shapes and so on.
            if (((child.getNodeType()) == (NodeType.RUN)))
            {
                // Say we found the node that we want, do something useful.
                Run run = (Run) child;
                msConsole.writeLine(run.getText());
            }
        }

        //ExEnd
    }

    //ExStart
    //ExFor:Node.NextSibling
    //ExFor:CompositeNode.FirstChild
    //ExFor:Node.IsComposite
    //ExFor:CompositeNode.IsComposite
    //ExFor:Node.NodeTypeToString
    //ExSummary:Shows how to efficiently visit all direct and indirect children of a composite node.
    @Test //ExSkip
    public void recurseAllNodes() throws Exception
    {
        // Open a document.
        Document doc = new Document(getMyDir() + "Node.RecurseAllNodes.doc");

        // Invoke the recursive function that will walk the tree.
        traverseAllNodes(doc);
    }

    /// <summary>
    /// A simple function that will walk through all children of a specified node recursively 
    /// and print the type of each node to the screen.
    /// </summary>
    @Test (enabled = false)
    public void traverseAllNodes(CompositeNode parentNode)
    {
        // This is the most efficient way to loop through immediate children of a node.
        for (Node childNode = parentNode.getFirstChild(); childNode != null; childNode = childNode.getNextSibling())
        {
            // Do some useful work.
            msConsole.writeLine(Node.nodeTypeToString(childNode.getNodeType()));

            // Recurse into the node if it is a composite node.
            if (childNode.isComposite())
                traverseAllNodes((CompositeNode) childNode);
        }
    }
    //ExEnd

    @Test
    public void removeNodes() throws Exception
    {
        Document doc = new Document();

        //ExStart
        //ExFor:Node
        //ExFor:Node.NodeType
        //ExFor:Node.Remove
        //ExSummary:Shows how to remove all nodes of a specific type from a composite node. In this example we remove tables from a section body.
        // Get the section that we want to work on.
        Section section = doc.getSections().get(0);
        Body body = section.getBody();

        // Select the first child node in the body.
        Node curNode = body.getFirstChild();

        while (curNode != null)
        {
            // Save the pointer to the next sibling node because if the current 
            // node is removed from the parent in the next step, we will have 
            // no way of finding the next node to continue the loop.
            Node nextNode = curNode.getNextSibling();

            // A section body can contain Paragraph and Table nodes.
            // If the node is a Table, remove it from the parent.
            if (((curNode.getNodeType()) == (NodeType.TABLE)))
                curNode.remove();

            // Continue going through child nodes until null (no more siblings) is reached.
            curNode = nextNode;
        }

        //ExEnd
    }

    @Test
    public void enumNextSibling() throws Exception
    {
        Document doc = new Document();

        //ExStart
        //ExFor:CompositeNode.FirstChild
        //ExFor:Node.NextSibling
        //ExFor:Node.NodeTypeToString
        //ExFor:Node.NodeType
        //ExSummary:Shows how to enumerate immediate child nodes of a composite node using NextSibling. In this example we enumerate all paragraphs of a section body.
        // Get the section that we want to work on.
        Section section = doc.getSections().get(0);
        Body body = section.getBody();

        // Loop starting from the first child until we reach null.
        for (Node node = body.getFirstChild(); node != null; node = node.getNextSibling())
        {
            // Output the types of the nodes that we come across.
            msConsole.writeLine(Node.nodeTypeToString(node.getNodeType()));
        }

        //ExEnd
    }

    @Test
    public void typedAccess() throws Exception
    {
        Document doc = new Document();

        //ExStart
        //ExFor:Story.Tables
        //ExFor:Table.FirstRow
        //ExFor:Table.LastRow
        //ExFor:TableCollection
        //ExSummary:Demonstrates how to use typed properties to access nodes of the document tree.
        // Quick typed access to the first child Section node of the Document.
        Section section = doc.getFirstSection();

        // Quick typed access to the Body child node of the Section.
        Body body = section.getBody();

        // Quick typed access to all Table child nodes contained in the Body.
        TableCollection tables = body.getTables();

        for (Table table : tables.<Table>OfType() !!Autoporter error: Undefined expression type )
        {
            // Quick typed access to the first row of the table.
            table.getFirstRow()?.Remove();

            // Quick typed access to the last row of the table.
            table.getLastRow()?.Remove();
        }

        //ExEnd
    }

    @Test
    public void updateFieldsInRange() throws Exception
    {
        Document doc = new Document();

        //ExStart
        //ExFor:Range.UpdateFields
        //ExSummary:Demonstrates how to update document fields in the body of the first section only.
        doc.getFirstSection().getBody().getRange().updateFields();
        //ExEnd
    }

    @Test
    public void removeChild() throws Exception
    {
        Document doc = new Document();

        //ExStart
        //ExFor:CompositeNode.LastChild
        //ExFor:Node.PreviousSibling
        //ExFor:CompositeNode.RemoveChild
        //ExSummary:Demonstrates use of methods of Node and CompositeNode to remove a section before the last section in the document.
        // Document is a CompositeNode and LastChild returns the last child node in the Document node.
        // Since the Document can contain only Section nodes, the last child is the last section.
        Node lastSection = doc.getLastChild();

        // Each node knows its next and previous sibling nodes.
        // Previous sibling of a section is a section before the specified section.
        // If the node is the first child, PreviousSibling will return null.
        Node sectionBeforeLast = lastSection.getPreviousSibling();

        if (sectionBeforeLast != null)
            doc.removeChild(sectionBeforeLast);
        //ExEnd
    }

    @Test
    public void compositeNode_SelectNodes() throws Exception
    {
        //ExStart
        //ExFor:CompositeNode.SelectSingleNode
        //ExFor:CompositeNode.SelectNodes
        //ExFor:NodeList.GetEnumerator
        //ExFor:NodeList.ToArray
        //ExSummary:Shows how to select certain nodes by using an XPath expression.
        Document doc = new Document(getMyDir() + "Table.Document.doc");

        // This expression will extract all paragraph nodes which are descendants of any table node in the document.
        // This will return any paragraphs which are in a table.
        NodeList nodeList = doc.selectNodes("//Table//Paragraph");

        // Iterate through the list with an enumerator and print the contents of every paragraph in each cell of the table
        int index = 0;
        Iterator<Node> e = nodeList.iterator();
        try /*JAVA: was using*/
        {
            while (e.hasNext())
            {
                msConsole.writeLine($"Table paragraph index {index++}, contents: \"{e.Current.GetText().Trim()}\"");
            }
        }
        finally { if (e != null) e.close(); }

        // This expression will select any paragraphs that are direct children of any body node in the document.
        nodeList = doc.selectNodes("//Body/Paragraph");

        // We can treat the list as an array too
        msAssert.areEqual(4, nodeList.toArray().length);

        // Use SelectSingleNode to select the first result of the same expression as above.
        Node node = doc.selectSingleNode("//Body/Paragraph");
        //ExEnd
    }

    @Test
    public void testNodeIsInsideField() throws Exception
    {
        //ExStart:
        //ExFor:CompositeNode.SelectNodes
        //ExFor:CompositeNode.GetChild
        //ExSummary:Shows how to test if a node is inside a field by using an XPath expression.
        // Let's pick a document we know has some fields in.
        Document doc = new Document(getMyDir() + "MailMerge.MergeImage.doc");

        // Let's say we want to check if the Run below is inside a field.
        Run run = (Run) doc.getChild(NodeType.RUN, 5, true);

        // Evaluate the XPath expression. The resulting NodeList will contain all nodes found inside a field a field (between FieldStart 
        // and FieldEnd exclusive). There can however be FieldStart and FieldEnd nodes in the list if there are nested fields 
        // in the path. Currently does not find rare fields in which the FieldCode or FieldResult spans across multiple paragraphs.
        NodeList resultList =
            doc.selectNodes("//FieldStart/following-sibling::node()[following-sibling::FieldEnd]");

        // Check if the specified run is one of the nodes that are inside the field.
        for (Node node : resultList)
        {
            if (node == run)
            {
                msConsole.writeLine("The node is found inside a field");
                break;
            }
        }

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
        Document doc = new Document(getMyDir() + "Document.doc");

        // Remove smart tags from the first paragraph in the document.
        doc.getFirstSection().getBody().getFirstParagraph().removeSmartTags();
        //ExEnd
    }

    @Test
    public void getIndexOfNode() throws Exception
    {
        //ExStart
        //ExFor:CompositeNode.IndexOf
        //ExSummary:Shows how to get the index of a given child node from its parent.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Get the body of the first section in the document.
        Body body = doc.getFirstSection().getBody();
        // Retrieve the index of the last paragraph in the body.
        int index = body.getChildNodes().indexOf(body.getLastParagraph());
        //ExEnd

        // Verify that the index is correct.
        msAssert.areEqual(24, index);
    }

    @Test
    public void getNodeTypeEnums() throws Exception
    {
        //ExStart
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
        //ExSummary:Shows how to retrieve the NodeType enumeration of nodes.
        Document doc = new Document(getMyDir() + "Document.doc");

        // Let's pick a node that we can't be quite sure of what type it is.
        // In this case lets pick the first node of the first paragraph in the body of the document
        Node node = doc.getFirstSection().getBody().getFirstParagraph().getFirstChild();
        msConsole.writeLine("NodeType of first child: " + Node.nodeTypeToString(node.getNodeType()));

        // This time let's pick a node that we know the type of. Create a new paragraph and a table node.
        Paragraph para = new Paragraph(doc);
        Table table = new Table(doc);

        // Access to NodeType for typed nodes will always return their specific NodeType. 
        // i.e A paragraph node will always return NodeType.Paragraph, a table node will always return NodeType.Table.
        msConsole.writeLine("NodeType of Paragraph: " + Node.nodeTypeToString(para.getNodeType()));
        msConsole.writeLine("NodeType of Table: " + Node.nodeTypeToString(table.getNodeType()));
        //ExEnd
    }

    @Test
    public void convertNodeToHtmlWithDefaultOptions() throws Exception
    {
        //ExStart
        //ExFor:Node.ToString(SaveFormat)
        //ExSummary:Exports the content of a node to String in HTML format using default options.
        Document doc = new Document(getMyDir() + "Document.doc");

        // Extract the last paragraph in the document to convert to HTML.
        Node node = doc.getLastSection().getBody().getLastParagraph();

        // When ToString is called using the SaveFormat overload then conversion is executed using default save options. 
        // When saving to HTML using default options the following settings are set:
        //   ExportImagesAsBase64 = true
        //   CssStyleSheetType = CssStyleSheetType.Inline
        //   ExportFontResources = false
        String nodeAsHtml = node.toString(SaveFormat.HTML);
        //ExEnd

        msAssert.areEqual(
            "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:12pt\"><span style=\"font-family:'Times New Roman'\">Hello World!</span></p>",
            nodeAsHtml);
    }

    @Test
    public void convertNodeToHtmlWithSaveOptions() throws Exception
    {
        //ExStart
        //ExFor:Node.ToString(SaveOptions)
        //ExSummary:Exports the content of a node to String in HTML format using custom specified options.
        Document doc = new Document(getMyDir() + "Document.doc");

        // Extract the last paragraph in the document to convert to HTML.
        Node node = doc.getLastSection().getBody().getLastParagraph();

        // Create an instance of HtmlSaveOptions and set a few options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setExportHeadersFootersMode(ExportHeadersFootersMode.PER_SECTION);
            saveOptions.setExportRelativeFontSize(true);
        }

        // Convert the document to HTML and return as a String. Pass the instance of HtmlSaveOptions to
        // to use the specified options during the conversion.
        String nodeAsHtml = node.toString(saveOptions);
        //ExEnd

        msAssert.areEqual(
            "<p style=\"margin-top:0pt; margin-bottom:0pt\"><span style=\"font-family:'Times New Roman'\">Hello World!</span></p>",
            nodeAsHtml);
    }

    @Test
    public void typedNodeCollectionToArray() throws Exception
    {
        Document doc = new Document();

        //ExStart
        //ExFor:ParagraphCollection.ToArray
        //ExSummary:Demonstrates typed implementations of ToArray on classes derived from NodeCollection.
        // You can use ToArray to return a typed array of nodes.
        Paragraph[] paras = doc.getFirstSection().getBody().getParagraphs().toArray();
        //ExEnd

        Assert.That(paras.length, Is.GreaterThan(0));
    }

    @Test
    public void nodeEnumerationHotRemove() throws Exception
    {
        //ExStart
        //ExFor:ParagraphCollection.ToArray
        //ExSummary:Demonstrates how to use "hot remove" to remove a node during enumeration.
        DocumentBuilder builder = new DocumentBuilder();
        builder.writeln("The first paragraph");
        builder.writeln("The second paragraph");
        builder.writeln("The third paragraph");
        builder.writeln("The fourth paragraph");

        // Hot remove allows a node to be removed from a live collection and have the enumeration continue.
        for (Paragraph para : builder.getDocument().getFirstSection().getBody().getChildNodes(NodeType.PARAGRAPH, true)
            .<Paragraph>OfType() !!Autoporter error: Undefined expression type )
        {
            if (para.getRange().getText().contains("third"))
            {
                // Enumeration will continue even after this node is removed.
                para.remove();
            }
        }

        //ExEnd
    }

    @Test
    public void enumerationHotRemoveLimitations() throws Exception
    {
        //ExStart
        //ExFor:ParagraphCollection.ToArray
        //ExSummary:Demonstrates an example breakage of the node collection enumerator.
        DocumentBuilder builder = new DocumentBuilder();
        builder.writeln("The first paragraph");
        builder.writeln("The second paragraph");
        builder.writeln("The third paragraph");
        builder.writeln("The fourth paragraph");

        // This causes unexpected behavior, the fourth paragraph in the collection is not visited.
        for (Paragraph para : builder.getDocument().getFirstSection().getBody().getChildNodes(NodeType.PARAGRAPH, true)
            .<Paragraph>OfType() !!Autoporter error: Undefined expression type )
        {
            if (para.getRange().getText().contains("third"))
            {
                para.getPreviousSibling().remove();
                para.remove();
            }
        }
        //ExEnd
    }

    @Test
    public void compositeNodeChildren() throws Exception
    {
        //ExStart
        //ExFor:CompositeNode.Count
        //ExFor:CompositeNode.GetChildNodes(NodeType[], Boolean)
        //ExFor:CompositeNode.InsertAfter(Node, Node)
        //ExFor:CompositeNode.InsertBefore(Node, Node)
        //ExFor:CompositeNode.PrependChild(Node) 
        //ExFor:Paragraph.GetText
        //ExSummary:Shows how to add, update and delete child nodes from within a CompositeNode.
        Document doc = new Document();

        // An empty document has one paragraph by default
        msAssert.areEqual(1, doc.getFirstSection().getBody().getParagraphs().getCount());

        // A paragraph is a composite node because it can contain runs, which are another type of node
        Paragraph paragraph = doc.getFirstSection().getBody().getFirstParagraph();
        Run paragraphText = new Run(doc, "Initial text. ");
        paragraph.appendChild(paragraphText);

        // We will place these 3 children into the main text of our paragraph
        Run run1 = new Run(doc, "Run 1. ");
        Run run2 = new Run(doc, "Run 2. ");
        Run run3 = new Run(doc, "Run 3. ");

        // We initialized them but not in our paragraph yet
        msAssert.areEqual("Initial text. " + (char) 12, paragraph.getText());

        // Insert run2 before initial paragraph text. This will be at the start of the paragraph
        paragraph.insertBefore(run2, paragraphText);

        // Insert run3 after initial paragraph text. This will be at the end of the paragraph
        paragraph.insertAfter(run3, paragraphText);

        // Insert run1 before every other child node. run2 was the start of the paragraph, now it will be run1
        paragraph.prependChild(run1);

        msAssert.areEqual("Run 1. Run 2. Initial text. Run 3. " + (char) 12, paragraph.getText());
        msAssert.areEqual(4, paragraph.getChildNodes(NodeType.ANY, true).getCount());

        // Access the child node collection and update/delete children
        ((Run) paragraph.getChildNodes(NodeType.RUN, true).get(1)).setText("Updated run 2. ");
        paragraph.getChildNodes(NodeType.RUN, true).remove(paragraphText);

        msAssert.areEqual("Run 1. Updated run 2. Run 3. " + (char) 12, paragraph.getText());
        msAssert.areEqual(3, paragraph.getChildNodes(NodeType.ANY, true).getCount());
        //ExEnd
    }

    //ExStart
    //ExFor:Aspose.Words.CompositeNode.CreateNavigator
    //ExSummary:Shows how to create an XPathNavigator and use it to traverse and read nodes.
    @Test //ExSkip
    public void nodeXPathNavigator() throws Exception
    {
        // Create a blank document
        Document doc = new Document();

        // A document is a composite node so we can make a navigator straight away
        XPathNavigator navigator = doc.CreateNavigator();

        // Our root is the document node with 1 child, which is the first section
        if (navigator != null)
        {
            msAssert.areEqual("Document", navigator.Name);
            msAssert.areEqual(false, navigator.MoveToNext());
            msAssert.areEqual(1, navigator.SelectChildren(XPathNodeType.All).Count);

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
        }
    }

    /// <summary>
    /// This will traverse all children of a composite node and map the structure in the style of a directory tree.
    /// Amount of space indentation indicates depth relative to initial node. Only runs will have their values printed.
    /// </summary>
    private void mapDocument(XPathNavigator navigator, StringBuilder stringBuilder, int depth)
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

        builder.insertImage(BitmapPal.loadNativeImage(getImageDir() + "Aspose.Words.gif"));
        builder.getCurrentParagraph().getParentNode().removeAllChildren();
    }

    /// <summary>
    /// Prints all inserted/removed nodes as well as their parent nodes
    /// </summary>
    private static class NodeChangingPrinter implements INodeChangingCallback
    {
        public void /*INodeChangingCallback.*/nodeInserting(NodeChangingArgs args)
        {
            msAssert.areEqual(NodeChangingAction.INSERT, args.getAction());
            msAssert.areEqual(null, args.getOldParent());
        }

        public void /*INodeChangingCallback.*/nodeInserted(NodeChangingArgs args)
        {
            msAssert.areEqual(NodeChangingAction.INSERT, args.getAction());
            Assert.assertNotNull(args.getNewParent());

            msConsole.writeLine($"Inserted node:");
            msConsole.writeLine($"\tType:\t{args.Node.NodeType}");

            if (!"".equals(msString.trim(args.getNode().getText())))
            {
                msConsole.writeLine($"\tText:\t\"{args.Node.GetText().Trim()}\"");
            }

            msConsole.writeLine($"\tHash:\t{args.Node.GetHashCode()}");
            msConsole.writeLine($"\tParent:\t{args.NewParent.NodeType} ({args.NewParent.GetHashCode()})");
        }

        public void /*INodeChangingCallback.*/nodeRemoving(NodeChangingArgs args)
        {
            msAssert.areEqual(NodeChangingAction.REMOVE, args.getAction());
        }

        public void /*INodeChangingCallback.*/nodeRemoved(NodeChangingArgs args)
        {
            msAssert.areEqual(NodeChangingAction.REMOVE, args.getAction());
            Assert.assertNull(args.getNewParent());

            msConsole.writeLine($"Removed node: {args.Node.NodeType} ({args.Node.GetHashCode()})");
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
        msAssert.areEqual(2, runs.getCount());

        // We can insert a node into the RunCollection manually to achieve the same effect
        Run newRun = new Run(doc, "Run 3. ");
        runs.insert(3, newRun);

        Assert.assertTrue(runs.contains(newRun));
        msAssert.areEqual("Run 1. Run 2. Run 3.", msString.trim(doc.getText()));

        // Text can also be deleted from the document by accessing individual Runs via the RunCollection and editing or removing them
        Run run = runs.get(1);
        runs.remove(run);
        msAssert.areEqual("Run 1. Run 3.", msString.trim(doc.getText()));

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

        builder.insertImage(BitmapPal.loadNativeImage(getImageDir() + "Aspose.Words.gif"));
        // Get all run nodes, of which we put 3 in the entire document
        NodeList nodeList = doc.selectNodes("//Run");
        msAssert.areEqual(3, nodeList.getCount());

        // Using a double forward slash, select all Run nodes that are indirect descendants of a Table node,
        // which would in this case be the runs inside the two cells we inserted
        nodeList = doc.selectNodes("//Table//Run");
        msAssert.areEqual(2, nodeList.getCount());

        // Single forward slashes specify direct descendant relationships,
        // of which we skipped quite a few by using double slashes
        msAssert.areEqual(doc.selectNodes("//Table//Run"), doc.selectNodes("//Table/Row/Cell/Paragraph/Run"));

        // We can access the actual nodes via a NodeList too
        nodeList = doc.selectNodes("//Shape");
        msAssert.areEqual(1, nodeList.getCount());
        Shape shape = (Shape)nodeList.get(0);
        Assert.assertTrue(shape.hasImage());
        //ExEnd
    }
}
