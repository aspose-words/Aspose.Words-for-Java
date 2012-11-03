//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package Examples;

import com.aspose.words.*;
import org.testng.annotations.Test;
import org.testng.Assert;


public class ExNode extends ExBase
{
    @Test
    public void useNodeType() throws Exception
    {
        //ExStart
        //ExFor:NodeType
        //ExId:UseNodeType
        //ExSummary:The following example shows how to use the NodeType enumeration.
        Document doc = new Document();

        // Returns NodeType.Document
        int type = doc.getNodeType();
        //ExEnd
    }

	@Test
	public void cloneNode() throws Exception
	{
		//ExStart
		//ExFor:Node
		//ExFor:Node.Clone
		//ExSummary:Shows how to clone nodes with and without their child nodes.
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

		Assert.assertTrue(((CompositeNode)cloneWithChildren).hasChildNodes());
		Assert.assertFalse(((CompositeNode)cloneWithoutChildren).hasChildNodes());
	}

	@Test
    public void getParentNode() throws Exception
    {
        //ExStart
        //ExFor:Node.ParentNode
        //ExId:AccessParentNode
        //ExSummary:Shows how to access the parent node.
        // Create a new empty document. It has one section.
        Document doc = new Document();

        // The section is the first child node of the document.
        Node section = doc.getFirstChild();

        // The section's parent node is the document.
        System.out.println("Section parent is the document: " + (doc == section.getParentNode()));
        //ExEnd

        Assert.assertEquals(doc, section.getParentNode());
    }

    @Test
    public void ownerDocument() throws Exception
    {
        //ExStart
        //ExFor:Node.Document
        //ExFor:Node.ParentNode
        //ExId:CreatingNodeRequiresDocument
        //ExSummary:Shows that when you create any node, it requires a document that will own the node.
        // Open a file from disk.
        Document doc = new Document();

        // Creating a new node of any type requires a document passed into the constructor.
        Paragraph para = new Paragraph(doc);

        // The new paragraph node does not yet have a parent.
        System.out.println("Paragraph has no parent node: " + (para.getParentNode() == null));

        // But the paragraph node knows its document.
        System.out.println("Both nodes' documents are the same: " + (para.getDocument() == doc));

        // The fact that a node always belongs to a document allows us to access and modify
        // properties that reference the document-wide data such as styles or lists.
        para.getParagraphFormat().setStyleName("Heading 1");

        // Now add the paragraph to the main text of the first section.
        doc.getFirstSection().getBody().appendChild(para);

        // The paragraph node is now a child of the Body node.
        System.out.println("Paragraph has a parent node: " + (para.getParentNode() != null));
        //ExEnd

        Assert.assertEquals(doc, para.getDocument());
        Assert.assertNotNull(para.getParentNode());
    }

    @Test
    public void enumerateChildNodes() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:Node
        //ExFor:CompositeNode
        //ExFor:CompositeNode.GetChild
        //ExSummary:Shows how to extract a specific child node from a CompositeNode by using the GetChild method and passing the NodeType and index.
        Paragraph paragraph = (Paragraph)doc.getChild(NodeType.PARAGRAPH, 0, true);
        //ExEnd

        //ExStart
        //ExFor:CompositeNode.ChildNodes
        //ExFor:CompositeNode.GetEnumerator
        //ExId:ChildNodesForEach
        //ExSummary:Shows how to enumerate immediate children of a CompositeNode using the enumerator provided by the ChildNodes collection.
        NodeCollection children = paragraph.getChildNodes();
        for (Node child : (Iterable<Node>) children)
        {
            // Paragraph may contain children of various types such as runs, shapes and so on.
            if (child.getNodeType() == NodeType.RUN)
            {
                // Say we found the node that we want, do something useful.
                Run run = (Run)child;
                System.out.println(run.getText());
            }
        }
        //ExEnd
    }

    @Test
    public void indexChildNodes() throws Exception
    {
        Document doc = new Document();
        Paragraph paragraph = (Paragraph)doc.getChild(NodeType.PARAGRAPH, 0, true);

        //ExStart
        //ExFor:NodeCollection.Count
        //ExFor:NodeCollection.Item
        //ExId:ChildNodesIndexer
        //ExSummary:Shows how to enumerate immediate children of a CompositeNode using indexed access.
        NodeCollection children = paragraph.getChildNodes();
        for (int i = 0; i < children.getCount(); i++)
        {
            Node child = children.get(i);

            // Paragraph may contain children of various types such as runs, shapes and so on.
            if (child.getNodeType() == NodeType.RUN)
            {
                // Say we found the node that we want, do something useful.
                Run run = (Run)child;
                System.out.println(run.getText());
            }
        }
        //ExEnd
    }

    @Test
    public void recurseAllNodesCaller() throws Exception
    {
        recurseAllNodes();
    }

    //ExStart
    //ExFor:Node.NextSibling
    //ExFor:CompositeNode.FirstChild
    //ExFor:Node.IsComposite
    //ExFor:CompositeNode.IsComposite
    //ExFor:Node.NodeTypeToString
    //ExId:RecurseAllNodes
    //ExSummary:Shows how to efficiently visit all direct and indirect children of a composite node.
    public void recurseAllNodes() throws Exception
    {
        // Open a document.
        Document doc = new Document(getMyDir() + "Node.RecurseAllNodes.doc");

        // Invoke the recursive function that will walk the tree.
        traverseAllNodes(doc);
    }

    /**
     * A simple function that will walk through all children of a specified node recursively
     * and print the type of each node to the screen.
     */
    public void traverseAllNodes(CompositeNode parentNode) throws Exception
    {
        // This is the most efficient way to loop through immediate children of a node.
        for (Node childNode = parentNode.getFirstChild(); childNode != null; childNode = childNode.getNextSibling())
        {
            // Do some useful work.
            System.out.println(Node.nodeTypeToString(childNode.getNodeType()));

            // Recurse into the node if it is a composite node.
            if (childNode.isComposite())
                traverseAllNodes((CompositeNode)childNode);
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
            if (curNode.getNodeType() == NodeType.TABLE)
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
            System.out.println(Node.nodeTypeToString(node.getNodeType()));
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
        //ExId:TypedPropertiesAccess
        //ExSummary:Demonstrates how to use typed properties to access nodes of the document tree.
        // Quick typed access to the first child Section node of the Document.
        Section section = doc.getFirstSection();

        // Quick typed access to the Body child node of the Section.
        Body body = section.getBody();

        // Quick typed access to all Table child nodes contained in the Body.
        TableCollection tables = body.getTables();

        for (Table table : tables)
        {
            // Quick typed access to the first row of the table.
            if (table.getFirstRow() != null)
                table.getFirstRow().remove();

            // Quick typed access to the last row of the table.
            if (table.getLastRow() != null)
                table.getLastRow().remove();
        }
        //ExEnd
    }

    @Test
    public void UpdateFieldsInRange() throws Exception
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
        //ExSummary:Shows how to select certain nodes using an XPath expression.
        Document doc = new Document(getMyDir() + "Table.Document.doc");

        // This expression will extract all paragraph nodes which are descendants of any table node in the document.
        // This will return any paragraphs which are in a table.
        NodeList nodeList = doc.selectNodes("//Table//Paragraph");

        // This expression will select any paragraphs that are direct children of any body node in the document.
        nodeList = doc.selectNodes("//Body/Paragraph");

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
        //ExSummary:Shows how to test if a node is inside a field using an XPath expression.
        // Let's pick a document we know has some fields in.
        Document doc = new Document(getMyDir() + "MailMerge.MergeImage.doc");

        // Let's say we want to check if the Run below is inside a field.
        Run run = (Run)doc.getChild(NodeType.RUN, 5, true);

        // Evaluate the XPath expression. The resulting NodeList will contain all nodes found inside a field a field (between FieldStart
        // and FieldEnd exclusive). There can however be FieldStart and FieldEnd nodes in the list if there are nested fields
        // in the path. Currently does not find rare fields in which the FieldCode or FieldResult spans across multiple paragraphs.
        NodeList resultList = doc.selectNodes("//FieldStart/following-sibling::node()[following-sibling::FieldEnd]");

        // Check if the specified run is one of the nodes that are inside the field.
        for (Node node : (Iterable<Node>)resultList)
        {
            if (node == run)
            {
                System.out.println("The node is found inside a field");
                break;
            }
        }
        //ExEnd
    }

    @Test
    public void createAndAddParagraphNode() throws Exception
    {
        //ExStart
        //ExId:CreateAndAddParagraphNode
        //ExSummary:Creates and adds a paragraph node.
        Document doc = new Document();

        Paragraph para = new Paragraph(doc);

        Section section = doc.getLastSection();
        section.getBody().appendChild(para);
        //ExEnd
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
    public void GetIndexOfNode() throws Exception
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
        Assert.assertEquals(index, 24);
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
        //ExFor:DrawingML.NodeType
        //ExId:GetNodeTypeEnums
        //ExSummary:Shows how to retrieve the NodeType enumeration of nodes.
        Document doc = new Document(getMyDir() + "Document.doc");

        // Let's pick a node that we can't be quite sure of what type it is.
        // In this case lets pick the first node of the first paragraph in the body of the document
        Node node = doc.getFirstSection().getBody().getFirstParagraph().getFirstChild();
        System.out.println("NodeType of first child: " + Node.nodeTypeToString(node.getNodeType()));

        // This time let's pick a node that we know the type of. Create a new paragraph and a table node.
        Paragraph para = new Paragraph(doc);
        Table table = new Table(doc);

        // Access to NodeType for typed nodes will always return their specific NodeType.
        // i.e A paragraph node will always return NodeType.Paragraph, a table node will always return NodeType.Table.
        System.out.println("NodeType of Paragraph: " + Node.nodeTypeToString(para.getNodeType()));
        System.out.println("NodeType of Table: " + Node.nodeTypeToString(table.getNodeType()));
        //ExEnd
    }

    @Test
    public void convertNodeToHtmlWithDefaultOptions() throws Exception
    {
        //ExStart
        //ExFor:Node.ToString(SaveFormat)
        //ExSummary:Exports the content of a node to string in HTML format using default options.
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

        Assert.assertEquals(nodeAsHtml, "<p style=\"margin:0pt\"><span style=\"font-family:'Times New Roman'; font-size:12pt\">Hello World!</span></p>");
    }

    @Test
    public void convertNodeToHtmlWithSaveOptions() throws Exception
    {
        //ExStart
        //ExFor:Node.ToString(SaveOptions)
        //ExSummary:Exports the content of a node to string in HTML format using custom specified options.
        Document doc = new Document(getMyDir() + "Document.doc");

        // Extract the last paragraph in the document to convert to HTML.
        Node node = doc.getLastSection().getBody().getLastParagraph();

        // Create an instance of HtmlSaveOptions and set a few options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportHeadersFootersMode(ExportHeadersFootersMode.PER_SECTION);
        saveOptions.setExportRelativeFontSize(true);

        // Convert the document to HTML and return as a string. Pass the instance of HtmlSaveOptions to
        // to use the specified options during the conversion.
        String nodeAsHtml = node.toString(saveOptions);
        //ExEnd

        Assert.assertEquals(nodeAsHtml, "<p style=\"margin:0pt\"><span style=\"font-family:'Times New Roman'\">Hello World!</span></p>");
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

        Assert.assertTrue(paras.length > 0);
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
        for (Paragraph para : (Iterable<Paragraph>)builder.getDocument().getFirstSection().getBody().getChildNodes(NodeType.PARAGRAPH, true))
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

        // This causes unexpected behavior, the fourth pargraph in the collection is not visited.
        for (Paragraph para : (Iterable<Paragraph>)builder.getDocument().getFirstSection().getBody().getChildNodes(NodeType.PARAGRAPH, true))
        {
            if (para.getRange().getText().contains("third"))
            {
                para.getPreviousSibling().remove();
                para.remove();
            }
        }
        //ExEnd
    }
}

