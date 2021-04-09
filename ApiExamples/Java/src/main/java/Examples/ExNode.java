package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.text.MessageFormat;
import java.util.Iterator;

public class ExNode extends ApiExampleBase {
    @Test
    public void cloneCompositeNode() throws Exception {
        //ExStart
        //ExFor:Node
        //ExFor:Node.Clone
        //ExSummary:Shows how to clone a composite node.
        Document doc = new Document();
        Paragraph para = doc.getFirstSection().getBody().getFirstParagraph();
        para.appendChild(new Run(doc, "Hello world!"));

        // Below are two ways of cloning a composite node.
        // 1 -  Create a clone of a node, and create a clone of each of its child nodes as well.
        Node cloneWithChildren = para.deepClone(true);

        Assert.assertTrue(((CompositeNode) cloneWithChildren).hasChildNodes());
        Assert.assertEquals("Hello world!", cloneWithChildren.getText().trim());

        // 2 -  Create a clone of a node just by itself without any children.
        Node cloneWithoutChildren = para.deepClone(false);

        Assert.assertFalse(((CompositeNode) cloneWithoutChildren).hasChildNodes());
        Assert.assertEquals("", cloneWithoutChildren.getText().trim());
        //ExEnd
    }

    @Test
    public void getParentNode() throws Exception {
        //ExStart
        //ExFor:Node.ParentNode
        //ExSummary:Shows how to access a node's parent node.
        Document doc = new Document();
        Paragraph para = doc.getFirstSection().getBody().getFirstParagraph();

        // Append a child Run node to the document's first paragraph.
        Run run = new Run(doc, "Hello world!");
        para.appendChild(run);

        // The paragraph is the parent node of the run node. We can trace this lineage
        // all the way to the document node, which is the root of the document's node tree.
        Assert.assertEquals(para, run.getParentNode());
        Assert.assertEquals(doc.getFirstSection().getBody(), para.getParentNode());
        Assert.assertEquals(doc.getFirstSection(), doc.getFirstSection().getBody().getParentNode());
        Assert.assertEquals(doc, doc.getFirstSection().getParentNode());
        //ExEnd
    }

    @Test
    public void ownerDocument() throws Exception {
        //ExStart
        //ExFor:Node.Document
        //ExFor:Node.ParentNode
        //ExSummary:Shows how to create a node and set its owning document.
        Document doc = new Document();
        Paragraph para = new Paragraph(doc);
        para.appendChild(new Run(doc, "Hello world!"));

        // We have not yet appended this paragraph as a child to any composite node.
        Assert.assertNull(para.getParentNode());

        // If a node is an appropriate child node type of another composite node,
        // we can attach it as a child only if both nodes have the same owner document.
        // The owner document is the document we passed to the node's constructor.
        // We have not attached this paragraph to the document, so the document does not contain its text.
        Assert.assertEquals(para.getDocument(), doc);
        Assert.assertEquals("", doc.getText().trim());

        // Since the document owns this paragraph, we can apply one of its styles to the paragraph's contents.
        para.getParagraphFormat().setStyleName("Heading 1");

        // Add this node to the document, and then verify its contents.
        doc.getFirstSection().getBody().appendChild(para);

        Assert.assertEquals(doc.getFirstSection().getBody(), para.getParentNode());
        Assert.assertEquals("Hello world!", doc.getText().trim());
        //ExEnd

        Assert.assertEquals(para.getDocument(), doc);
        Assert.assertNotNull(para.getParentNode());
    }

    @Test
    public void childNodesEnumerate() throws Exception {
        //ExStart
        //ExFor:Node
        //ExFor:Node.CustomNodeId
        //ExFor:NodeType
        //ExFor:CompositeNode
        //ExFor:CompositeNode.GetChild
        //ExFor:CompositeNode.ChildNodes
        //ExFor:CompositeNode.GetEnumerator
        //ExFor:NodeCollection.Count
        //ExFor:NodeCollection.Item
        //ExSummary:Shows how to traverse through a composite node's collection of child nodes.
        Document doc = new Document();

        // Add two runs and one shape as child nodes to the first paragraph of this document.
        Paragraph paragraph = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);
        paragraph.appendChild(new Run(doc, "Hello world! "));

        Shape shape = new Shape(doc, ShapeType.RECTANGLE);
        shape.setWidth(200.0);
        shape.setHeight(200.0);
        // Note that the 'CustomNodeId' is not saved to an output file and exists only during the node lifetime.
        shape.setCustomNodeId(100);
        shape.setWrapType(WrapType.INLINE);
        paragraph.appendChild(shape);

        paragraph.appendChild(new Run(doc, "Hello again!"));

        // Iterate through the paragraph's collection of immediate children,
        // and print any runs or shapes that we find within.
        NodeCollection children = paragraph.getChildNodes();

        Assert.assertEquals(3, paragraph.getChildNodes().getCount());

        for (Node child : (Iterable<Node>) children)
            switch (child.getNodeType()) {
                case NodeType.RUN:
                    System.out.println("Run contents:");
                    System.out.println("\t\"{child.GetText().Trim()}\"");
                    break;
                case NodeType.SHAPE:
                    Shape childShape = (Shape) child;
                    System.out.println("Shape:");
                    System.out.println("\t{childShape.ShapeType}, {childShape.Width}x{childShape.Height}");
                    Assert.assertEquals(100, shape.getCustomNodeId()); //ExSkip
                    break;
            }
        //ExEnd

        Assert.assertEquals(NodeType.RUN, paragraph.getChild(NodeType.RUN, 0, true).getNodeType());
        Assert.assertEquals("Hello world! Hello again!", doc.getText().trim());
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
    //ExSummary:Shows how to traverse a composite node's tree of child nodes.
    @Test //ExSkip
    public void recurseChildren() throws Exception {
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // Any node that can contain child nodes, such as the document itself, is composite.
        Assert.assertTrue(doc.isComposite());

        // Invoke the recursive function that will go through and print all the child nodes of a composite node.
        traverseAllNodes(doc, 0);
    }

    /// <summary>
    /// Recursively traverses a node tree while printing the type of each node
    /// with an indent depending on depth as well as the contents of all inline nodes.
    /// </summary>
    @Test(enabled = false)
    public void traverseAllNodes(CompositeNode parentNode, int depth) {
        for (Node childNode = parentNode.getFirstChild(); childNode != null; childNode = childNode.getNextSibling()) {
            System.out.println(MessageFormat.format("{0}{1}", String.format("	", depth), Node.nodeTypeToString(childNode.getNodeType())));

            // Recurse into the node if it is a composite node. Otherwise, print its contents if it is an inline node.
            if (childNode.isComposite()) {
                System.out.println();
                traverseAllNodes((CompositeNode) childNode, depth + 1);
            } else if (childNode instanceof Inline) {
                System.out.println(" - \"{childNode.GetText().Trim()}\"");
            } else {
                System.out.println();
            }
        }
    }
    //ExEnd

    @Test
    public void removeNodes() throws Exception {
        //ExStart
        //ExFor:Node
        //ExFor:Node.NodeType
        //ExFor:Node.Remove
        //ExSummary:Shows how to remove all child nodes of a specific type from a composite node.
        Document doc = new Document(getMyDir() + "Tables.docx");

        Assert.assertEquals(2, doc.getChildNodes(NodeType.TABLE, true).getCount());

        Node curNode = doc.getFirstSection().getBody().getFirstChild();

        while (curNode != null) {
            // Save the next sibling node as a variable in case we want to move to it after deleting this node.
            Node nextNode = curNode.getNextSibling();

            // A section body can contain Paragraph and Table nodes.
            // If the node is a Table, remove it from the parent.
            if (curNode.getNodeType() == NodeType.TABLE) {
                curNode.remove();
            }

            curNode = nextNode;
        }

        Assert.assertEquals(0, doc.getChildNodes(NodeType.TABLE, true).getCount());
        //ExEnd
    }

    @Test
    public void enumNextSibling() throws Exception {
        //ExStart
        //ExFor:CompositeNode.FirstChild
        //ExFor:Node.NextSibling
        //ExFor:Node.NodeTypeToString
        //ExFor:Node.NodeType
        //ExSummary:Shows how to use a node's NextSibling property to enumerate through its immediate children.
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        for (Node node = doc.getFirstSection().getBody().getFirstChild(); node != null; node = node.getNextSibling()) {
            System.out.println(Node.nodeTypeToString(node.getNodeType()));
        }
        //ExEnd
    }

    @Test
    public void typedAccess() throws Exception {
        //ExStart
        //ExFor:Story.Tables
        //ExFor:Table.FirstRow
        //ExFor:Table.LastRow
        //ExFor:TableCollection
        //ExSummary:Shows how to remove the first and last rows of all tables in a document.
        Document doc = new Document(getMyDir() + "Tables.docx");

        TableCollection tables = doc.getFirstSection().getBody().getTables();

        Assert.assertEquals(5, tables.get(0).getRows().getCount());
        Assert.assertEquals(4, tables.get(1).getRows().getCount());

        for (Table table : tables) {
            if (table.getFirstRow() != null) {
                table.getFirstRow().remove();
            }

            if (table.getLastRow() != null) {
                table.getLastRow().remove();
            }
        }

        Assert.assertEquals(3, tables.get(0).getRows().getCount());
        Assert.assertEquals(2, tables.get(1).getRows().getCount());
        //ExEnd
    }

    @Test
    public void removeChild() throws Exception {
        //ExStart
        //ExFor:CompositeNode.LastChild
        //ExFor:Node.PreviousSibling
        //ExFor:CompositeNode.RemoveChild
        //ExSummary:Shows how to use of methods of Node and CompositeNode to remove a section before the last section in the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Section 1 text.");
        builder.insertBreak(BreakType.SECTION_BREAK_CONTINUOUS);
        builder.writeln("Section 2 text.");

        // Both sections are siblings of each other.
        Section lastSection = (Section) doc.getLastChild();
        Section firstSection = (Section) lastSection.getPreviousSibling();

        // Remove a section based on its sibling relationship with another section.
        if (lastSection.getPreviousSibling() != null)
            doc.removeChild(firstSection);

        // The section we removed was the first one, leaving the document with only the second.
        Assert.assertEquals("Section 2 text.", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void selectCompositeNodes() throws Exception {
        //ExStart
        //ExFor:CompositeNode.SelectSingleNode
        //ExFor:CompositeNode.SelectNodes
        //ExFor:NodeList.GetEnumerator
        //ExFor:NodeList.ToArray
        //ExSummary:Shows how to select certain nodes by using an XPath expression.
        Document doc = new Document(getMyDir() + "Tables.docx");

        // This expression will extract all paragraph nodes,
        // which are descendants of any table node in the document.
        NodeList nodeList = doc.selectNodes("//Table//Paragraph");

        // Iterate through the list with an enumerator and print the contents of every paragraph in each cell of the table.
        int index = 0;

        Iterator<Node> e = nodeList.iterator();
        while (e.hasNext()) {
            Node currentNode = e.next();
            System.out.println(MessageFormat.format("Table paragraph index {0}, contents: \"{1}\"", index++, currentNode.getText().trim()));
        }

        // This expression will select any paragraphs that are direct children of any Body node in the document.
        nodeList = doc.selectNodes("//Body/Paragraph");

        // We can treat the list as an array.
        Assert.assertEquals(nodeList.toArray().length, 4);

        // Use SelectSingleNode to select the first result of the same expression as above.
        Node node = doc.selectSingleNode("//Body/Paragraph");

        Assert.assertEquals(Paragraph.class, node.getClass());
        //ExEnd
    }

    @Test
    public void testNodeIsInsideField() throws Exception {
        //ExStart
        //ExFor:CompositeNode.SelectNodes
        //ExSummary:Shows how to use an XPath expression to test whether a node is inside a field.
        Document doc = new Document(getMyDir() + "Mail merge destination - Northwind employees.docx");

        // The NodeList that results from this XPath expression will contain all nodes we find inside a field.
        // However, FieldStart and FieldEnd nodes can be on the list if there are nested fields in the path.
        // Currently does not find rare fields in which the FieldCode or FieldResult spans across multiple paragraphs.
        NodeList resultList =
                doc.selectNodes("//FieldStart/following-sibling::node()[following-sibling::FieldEnd]");

        // Check if the specified run is one of the nodes that are inside the field.
        System.out.println("Contents of the first Run node that's part of a field: {resultList.First(n => n.NodeType == NodeType.Run).GetText().Trim()}");
        //ExEnd
    }

    @Test
    public void createAndAddParagraphNode() throws Exception {
        Document doc = new Document();

        Paragraph para = new Paragraph(doc);

        Section section = doc.getLastSection();
        section.getBody().appendChild(para);
    }

    @Test
    public void removeSmartTagsFromCompositeNode() throws Exception {
        //ExStart
        //ExFor:CompositeNode.RemoveSmartTags
        //ExSummary:Removes all smart tags from descendant nodes of a composite node.
        Document doc = new Document(getMyDir() + "Smart tags.doc");

        Assert.assertEquals(8, doc.getChildNodes(NodeType.SMART_TAG, true).getCount());

        doc.removeSmartTags();

        Assert.assertEquals(0, doc.getChildNodes(NodeType.SMART_TAG, true).getCount());
        //ExEnd
    }

    @Test
    public void getIndexOfNode() throws Exception {
        //ExStart
        //ExFor:CompositeNode.IndexOf
        //ExSummary:Shows how to get the index of a given child node from its parent.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        Body body = doc.getFirstSection().getBody();

        // Retrieve the index of the last paragraph in the body of the first section.
        Assert.assertEquals(24, body.getChildNodes().indexOf(body.getLastParagraph()));
        //ExEnd
    }

    @Test
    public void convertNodeToHtmlWithDefaultOptions() throws Exception {
        //ExStart
        //ExFor:Node.ToString(SaveFormat)
        //ExFor:Node.ToString(SaveOptions)
        //ExSummary:Exports the content of a node to String in HTML format.
        Document doc = new Document(getMyDir() + "Document.docx");

        Node node = doc.getLastSection().getBody().getLastParagraph();

        // When we call the ToString method using the html SaveFormat overload,
        // it converts the node's contents to their raw html representation.
        Assert.assertEquals("<p style=\"margin-top:0pt; margin-bottom:8pt; line-height:108%; font-size:12pt\">" +
                "<span style=\"font-family:'Times New Roman'\">Hello World!</span>" +
                "</p>", node.toString(SaveFormat.HTML));

        // We can also modify the result of this conversion using a SaveOptions object.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportRelativeFontSize(true);

        Assert.assertEquals("<p style=\"margin-top:0pt; margin-bottom:8pt; line-height:108%\">" +
                "<span style=\"font-family:'Times New Roman'\">Hello World!</span>" +
                "</p>", node.toString(saveOptions));
        //ExEnd
    }

    @Test
    public void typedNodeCollectionToArray() throws Exception {
        //ExStart
        //ExFor:ParagraphCollection.ToArray
        //ExSummary:Shows how to create an array from a NodeCollection.
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        Paragraph[] paras = doc.getFirstSection().getBody().getParagraphs().toArray();

        Assert.assertEquals(22, paras.length);
        //ExEnd
    }

    @Test
    public void nodeEnumerationHotRemove() throws Exception {
        //ExStart
        //ExFor:ParagraphCollection.ToArray
        //ExSummary:Shows how to use "hot remove" to remove a node during enumeration.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("The first paragraph");
        builder.writeln("The second paragraph");
        builder.writeln("The third paragraph");
        builder.writeln("The fourth paragraph");

        // Remove a node from the collection in the middle of an enumeration.
        for (Paragraph para : doc.getFirstSection().getBody().getParagraphs().toArray())
            if (para.getRange().getText().contains("third"))
                para.remove();

        Assert.assertFalse(doc.getText().contains("The third paragraph"));
        //ExEnd
    }

    //ExStart
    //ExFor:NodeChangingAction
    //ExFor:NodeChangingArgs.Action
    //ExFor:NodeChangingArgs.NewParent
    //ExFor:NodeChangingArgs.OldParent
    //ExSummary:Shows how to use a NodeChangingCallback to monitor changes to the document tree in real-time as we edit it.
    @Test //ExSkip
    public void nodeChangingCallback() throws Exception {
        Document doc = new Document();
        doc.setNodeChangingCallback(new NodeChangingPrinter());

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");
        builder.startTable();
        builder.insertCell();
        builder.write("Cell 1");
        builder.insertCell();
        builder.write("Cell 2");
        builder.endTable();

        builder.insertImage(getImageDir() + "Logo.jpg");
        builder.getCurrentParagraph().getParentNode().removeAllChildren();
    }

    /// <summary>
    /// Prints every node insertion/removal as it takes place in the document.
    /// </summary>
    private static class NodeChangingPrinter implements INodeChangingCallback {
        public void nodeInserting(NodeChangingArgs args) {
            Assert.assertEquals(args.getAction(), NodeChangingAction.INSERT);
            Assert.assertEquals(args.getOldParent(), null);
        }

        public void nodeInserted(NodeChangingArgs args) {
            Assert.assertEquals(args.getAction(), NodeChangingAction.INSERT);
            Assert.assertNotNull(args.getNewParent());

            System.out.println("Inserted node:");
            System.out.println(MessageFormat.format("\tType:\t{0}", args.getNode().getNodeType()));

            if (!"".equals(args.getNode().getText().trim())) {
                System.out.println(MessageFormat.format("\tText:\t\"{0}\"", args.getNode().getText().trim()));
            }

            System.out.println(MessageFormat.format("\tHash:\t{0}", args.getNode().hashCode()));
            System.out.println(MessageFormat.format("\tParent:\t{0} ({1})", args.getNewParent().getNodeType(), args.getNewParent().hashCode()));
        }

        public void nodeRemoving(NodeChangingArgs args) {
            Assert.assertEquals(args.getAction(), NodeChangingAction.REMOVE);
        }

        public void nodeRemoved(NodeChangingArgs args) {
            Assert.assertEquals(args.getAction(), NodeChangingAction.REMOVE);
            Assert.assertNull(args.getNewParent());

            System.out.println(MessageFormat.format("Removed node: {0} ({1})", args.getNode().getNodeType(), args.getNode().hashCode()));
        }
    }
    //ExEnd

    @Test
    public void nodeCollection() throws Exception {
        //ExStart
        //ExFor:NodeCollection.Contains(Node)
        //ExFor:NodeCollection.Insert(Int32,Node)
        //ExFor:NodeCollection.Remove(Node)
        //ExSummary:Shows how to work with a NodeCollection.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add text to the document by inserting Runs using a DocumentBuilder.
        builder.write("Run 1. ");
        builder.write("Run 2. ");

        // Every invocation of the "Write()" method creates a new Run,
        // which then appears in the parent Paragraph's RunCollection.
        RunCollection runs = doc.getFirstSection().getBody().getFirstParagraph().getRuns();

        Assert.assertEquals(2, runs.getCount());

        // We can also insert a node into the RunCollection manually.
        Run newRun = new Run(doc, "Run 3. ");
        runs.insert(3, newRun);

        Assert.assertTrue(runs.contains(newRun));
        Assert.assertEquals("Run 1. Run 2. Run 3.", doc.getText().trim());

        // Access individual runs and remove them to remove their text from the document.
        Run run = runs.get(1);
        runs.remove(run);

        Assert.assertEquals("Run 1. Run 3.", doc.getText().trim());
        Assert.assertNotNull(run);
        Assert.assertFalse(runs.contains(run));
        //ExEnd
    }

    @Test
    public void nodeList() throws Exception {
        //ExStart
        //ExFor:NodeList.Count
        //ExFor:NodeList.Item(System.Int32)
        //ExSummary:Shows how to use XPaths to navigate a NodeList.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some nodes with a DocumentBuilder.
        builder.writeln("Hello world!");

        builder.startTable();
        builder.insertCell();
        builder.write("Cell 1");
        builder.insertCell();
        builder.write("Cell 2");
        builder.endTable();

        builder.insertImage(getImageDir() + "Logo.jpg");

        // Our document contains three Run nodes.
        NodeList runs = doc.selectNodes("//Run");

        Assert.assertEquals(3, runs.getCount());

        // Use a double forward slash to select all Run nodes
        // that are indirect descendants of a Table node, which would be the runs inside the two cells we inserted.
        runs = doc.selectNodes("//Table//Run");

        Assert.assertEquals(2, runs.getCount());

        // Single forward slashes specify direct descendant relationships,
        // which we skipped when we used double slashes.
        Assert.assertEquals(doc.selectNodes("//Table//Run"),
                doc.selectNodes("//Table/Row/Cell/Paragraph/Run"));

        // Access the shape that contains the image we inserted.
        NodeList shapes = doc.selectNodes("//Shape");

        Assert.assertEquals(1, shapes.getCount());

        Shape shape = (Shape) shapes.get(0);
        Assert.assertTrue(shape.hasImage());
        //ExEnd
    }
}
