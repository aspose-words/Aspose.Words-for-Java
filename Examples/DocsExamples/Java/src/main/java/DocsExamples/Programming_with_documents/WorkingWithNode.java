package DocsExamples.Programming_with_documents;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

@Test
public class WorkingWithNode extends DocsExamplesBase
{
    @Test
    public void useNodeType() throws Exception
    {
        //ExStart:UseNodeType            
        Document doc = new Document();
        
        int type = doc.getNodeType();
        //ExEnd:UseNodeType
    }

    @Test
    public void getParentNode() throws Exception
    {
        //ExStart:GetParentNode           
        Document doc = new Document();

        // The section is the first child node of the document.
        Node section = doc.getFirstChild();

        // The section's parent node is the document.
        System.out.println("Section parent is the document: " + (doc == section.getParentNode()));
        //ExEnd:GetParentNode           
    }

    @Test
    public void ownerDocument() throws Exception
    {
        //ExStart:OwnerDocument
        Document doc = new Document();

        // Creating a new node of any type requires a document passed into the constructor.
        Paragraph para = new Paragraph(doc);

        // The new paragraph node does not yet have a parent.
        System.out.println("Paragraph has no parent node: " + (para.getParentNode() == null));

        // But the paragraph node knows its document.
        System.out.println("Both nodes' documents are the same: " + (para.getDocument() == doc));

        // The fact that a node always belongs to a document allows us to access and modify
        // properties that reference the document-wide data, such as styles or lists.
        para.getParagraphFormat().setStyleName("Heading 1");

        // Now add the paragraph to the main text of the first section.
        doc.getFirstSection().getBody().appendChild(para);

        // The paragraph node is now a child of the Body node.
        System.out.println("Paragraph has a parent node: " + (para.getParentNode() != null));
        //ExEnd:OwnerDocument
    }

    @Test
    public void enumerateChildNodes() throws Exception
    {
        //ExStart:EnumerateChildNodes 
        Document doc = new Document();
        Paragraph paragraph = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        NodeCollection children = paragraph.getChildNodes();
        for (Node child : (Iterable<Node>) children)
        {
            // A paragraph may contain children of various types such as runs, shapes, and others.
            if (child.getNodeType() == NodeType.RUN)
            {
                Run run = (Run) child;
                System.out.println(run.getText());
            }
        }
        //ExEnd:EnumerateChildNodes
    }

    @Test
    //ExStart:RecurseAllNodes
    public void recurseAllNodes() throws Exception
    {
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // Invoke the recursive function that will walk the tree.
        traverseAllNodes(doc);
    }

    /// <summary>
    /// A simple function that will walk through all children of a specified node recursively 
    /// and print the type of each node to the screen.
    /// </summary>
    private void traverseAllNodes(CompositeNode parentNode)
    {
        // This is the most efficient way to loop through immediate children of a node.
        for (Node childNode = parentNode.getFirstChild(); childNode != null; childNode = childNode.getNextSibling())
        {
            System.out.println(Node.nodeTypeToString(childNode.getNodeType()));

            // Recurse into the node if it is a composite node.
            if (childNode.isComposite())
                traverseAllNodes((CompositeNode) childNode);
        }
    }
    //ExEnd:RecurseAllNodes

    @Test
    public void typedAccess() throws Exception
    {
        //ExStart:TypedAccess
        Document doc = new Document();

        Section section = doc.getFirstSection();
        Body body = section.getBody();
        
        // Quick typed access to all Table child nodes contained in the Body.
        TableCollection tables = body.getTables();

        for (Table table : tables)
        {
            // Quick typed access to the first row of the table.
            table.getFirstRow().remove();

            // Quick typed access to the last row of the table.
            table.getLastRow().remove();
        }
        //ExEnd:TypedAccess
    }

    @Test
    public void createAndAddParagraphNode() throws Exception
    {
        //ExStart:CreateAndAddParagraphNode
        Document doc = new Document();

        Paragraph para = new Paragraph(doc);

        Section section = doc.getLastSection();
        section.getBody().appendChild(para);
        //ExEnd:CreateAndAddParagraphNode
    }
}
