package DocsExamples.Complex_examples_and_helpers;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.ms.System.msConsole;
import com.aspose.words.Paragraph;
import com.aspose.words.LayoutEntityType;
import com.aspose.words.LayoutCollector;
import com.aspose.words.LayoutEnumerator;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import java.util.ArrayList;
import com.aspose.words.Row;
import com.aspose.ms.System.Collections.msDictionary;
import java.util.HashMap;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.ms.System.Collections.msArrayList;
import com.aspose.ms.System.Drawing.RectangleF;
import java.util.Iterator;


class DocumentLayoutHelper extends DocsExamplesBase
{
    @Test
    public void wrapperToAccessLayoutEntities() throws Exception
    {
        // This sample introduces the RenderedDocument class and other related classes which provide an API wrapper for 
        // the LayoutEnumerator. This allows you to access the layout entities of a document using a DOM style API.
        Document doc = new Document(getMyDir() + "Document layout.docx");

        RenderedDocument layoutDoc = new RenderedDocument(doc);

        // Get access to the line of the first page and print to the console.
        RenderedLine line = layoutDoc.Pages[0].Columns[0].Lines[2];
        System.out.println("Line: " + line.getText());

        // With a rendered line, the original paragraph in the document object model can be returned.
        Paragraph para = line.Paragraph;
        System.out.println("Paragraph text: " + para.getRange().getText());

        // Retrieve all the text that appears on the first page in plain text format (including headers and footers).
        String pageText = layoutDoc.Pages[0].Text;
        msConsole.writeLine();

        // Loop through each page in the document and print how many lines appear on each page.
        for (RenderedPage page : layoutDoc.Pages !!Autoporter error: Undefined expression type )
        {
            LayoutCollection<LayoutEntity> lines = page.getChildEntities(LayoutEntityType.LINE, true);
            msConsole.WriteLine("Page {0} has {1} lines.", page.PageIndex, lines.Count);
        }

        // This method provides a reverse lookup of layout entities for any given node
        // (except runs and nodes in the header and footer).
        msConsole.writeLine();
        System.out.println("The lines of the second paragraph:");
        for (RenderedLine paragraphLine : (Iterable<RenderedLine>) layoutDoc.getLayoutEntitiesOfNode(
            doc.getFirstSection().getBody().getParagraphs().get(1)))
        {
            System.out.println("\"{paragraphLine.Text.Trim()}\"");
            msConsole.WriteLine(paragraphLine.Rectangle.ToString());
            msConsole.writeLine();
        }
    }
}

/// <summary>
/// Provides an API wrapper for the LayoutEnumerator class to access the page layout
/// of a document presented in an object model like the design.
/// </summary>
public class RenderedDocument extends LayoutEntity
{
    /// <summary>
    /// Creates a new instance from the supplied Document class.
    /// </summary>
    /// <param name="doc">A document whose page layout model to enumerate.</param>
    /// <remarks><para>If page layout model of the document hasn't been built the enumerator calls
    /// <see cref="Document.UpdatePageLayout"/> to build it.</para>
    /// <para>Whenever document is updated and new page layout model is created,
    /// a new RenderedDocument instance must be used to access the changes.</para></remarks>
    public RenderedDocument(Document doc) throws Exception
    {
        mLayoutCollector = new LayoutCollector(doc);
        mEnumerator = new LayoutEnumerator(doc);

        processLayoutElements(this);
        linkLayoutMarkersToNodes(doc);
        collectLinesAndAddToMarkers();
    }

    /// <summary>
    /// Provides access to the pages of a document.
    /// </summary>
    public LayoutCollection<RenderedPage> Pages => private GetChildNodes<RenderedPage>getChildNodes();

    /// <summary>
    /// Returns all the layout entities of the specified node.
    /// </summary>
    /// <remarks>Note that this method does not work with Run nodes or nodes in the header or footer.</remarks>
    public LayoutCollection<LayoutEntity> getLayoutEntitiesOfNode(Node node)
    {
        if (!mLayoutCollector.getDocument().equals(node.getDocument()))
            throw new IllegalArgumentException("Node does not belong to the same document which was rendered.");

        if (node.getNodeType() == NodeType.DOCUMENT)
            return new LayoutCollection<LayoutEntity>(mChildEntities);

        ArrayList<LayoutEntity> entities = new ArrayList<LayoutEntity>();

        // Retrieve all entities from the layout document (inversion of LayoutEntityType.None).
        for (LayoutEntity entity : getChildEntities(~LayoutEntityType.NONE, true))
        {
            if (entity.ParentNode == node)
                entities.add(entity);

            // There is no table entity in rendered output, so manually check if rows belong to a table node.
            if (entity.Type == LayoutEntityType.ROW)
            {
                RenderedRow row = (RenderedRow) entity;
                if (row.Table == node)
                    entities.add(entity);
            }
        }

        return new LayoutCollection<LayoutEntity>(entities);
    }

    private void processLayoutElements(LayoutEntity current) throws Exception
    {
        do
        {
            LayoutEntity child = current.addChildEntity(mEnumerator);

            if (mEnumerator.moveFirstChild())
            {
                current = child;

                processLayoutElements(current);
                mEnumerator.moveParent();

                current = current.Parent;
            }
        } while (mEnumerator.moveNext());
    }

    private void collectLinesAndAddToMarkers()
    {
        collectLinesOfMarkersCore(LayoutEntityType.COLUMN);
        collectLinesOfMarkersCore(LayoutEntityType.COMMENT);
    }

    private void collectLinesOfMarkersCore(/*LayoutEntityType*/int type)
    {
        ArrayList<RenderedLine> collectedLines = new ArrayList<RenderedLine>();

        for (RenderedPage page : Pages !!Autoporter error: Undefined expression type )
        {
            for (LayoutEntity story : page.getChildEntities(type, false))
            {
                for (RenderedLine line : (Iterable<RenderedLine>) story.getChildEntities(LayoutEntityType.LINE, true))
                {
                    collectedLines.add(line);
                    for (RenderedSpan span : line.Spans !!Autoporter error: Undefined expression type )
                    {
                        if (mLayoutToNodeLookup.containsKey(span.getLayoutObject()))
                        {
                            if ("PARAGRAPH".equals(span.Kind) || "ROW".equals(span.Kind) || "CELL".equals(span.Kind) ||
                                "SECTION".equals(span.Kind))
                            {
                                Node node = mLayoutToNodeLookup.get(span.getLayoutObject());

                                if (node.getNodeType() == NodeType.ROW)
                                    node = ((Row) node).getLastCell().getLastParagraph();

                                for (RenderedLine collectedLine : collectedLines)
                                    collectedLine.setParentNode(node);

                                collectedLines = new ArrayList<RenderedLine>();
                            }
                            else
                            {
                                span.setParentNode(mLayoutToNodeLookup.get(span.getLayoutObject()));
                            }
                        }
                    }
                }
            }
        }
    }

    private void linkLayoutMarkersToNodes(Document doc) throws Exception
    {
        for (Node node : (Iterable<Node>) doc.getChildNodes(NodeType.ANY, true))
        {
            Object entity = mLayoutCollector.getEntity(node);

            if (entity != null)
                msDictionary.add(mLayoutToNodeLookup, entity, node);
        }
    }

    private /*final*/ LayoutCollector mLayoutCollector;
    private /*final*/ LayoutEnumerator mEnumerator;

    private /*final*/ HashMap<Object, Node> mLayoutToNodeLookup =
        new HashMap<Object, Node>();
}

/// <summary>
/// Provides the base class for rendered elements of a document.
/// </summary>
public abstract class LayoutEntity
{private mPageIndexmPageIndex;

    /// <summary>
    /// Returns bounding rectangle of the entity relative to the page top left corner (in points).
    /// </summary>
    public RectangleF Rectangle => private mRectanglemRectangle;private mTypemType;

    /// <summary>
    /// Exports the contents of the entity into a string in plain text format.
    /// </summary>
    public /*virtual*/ String getText()
    {
        StringBuilder builder = new StringBuilder();
        for (LayoutEntity entity : mChildEntities)
        {
            msStringBuilder.append(builder, entity.getText());
        }

        return builder.toString();
    }private mParentmParent;

    /// <summary>
    /// Returns the node that corresponds to this layout entity.  
    /// </summary>
    /// <remarks>This property may return null for spans that originate
    /// from Run nodes or nodes inside the header or footer.</remarks>
    public /*virtual*/ Node ParentNode => private mParentNodemParentNode;

    /// <summary>
    /// Internal method separate from ParentNode property to make code autoportable to VB.NET.
    /// </summary>
    /*virtual*/ void setParentNode(Node value)
    {
        mParentNode = value;
    }

    /// <summary>
    /// Reserved for internal use.
    /// </summary>
    Object getLayoutObject() { return mLayoutObject; }; void setLayoutObject(Object value) { mLayoutObject = value; };

    private Object mLayoutObject;

    /// <summary>
    /// Reserved for internal use.
    /// </summary>
    LayoutEntity addChildEntity(LayoutEnumerator it) throws Exception
    {
        LayoutEntity child = createLayoutEntityFromType(it);
        mChildEntities.add(child);

        return child;
    }

    private LayoutEntity createLayoutEntityFromType(LayoutEnumerator it) throws Exception
    {
        LayoutEntity childEntity;
        switch (it.getType())
        {
            case LayoutEntityType.CELL:
                childEntity = new RenderedCell();
                break;
            case LayoutEntityType.COLUMN:
                childEntity = new RenderedColumn();
                break;
            case LayoutEntityType.COMMENT:
                childEntity = new RenderedComment();
                break;
            case LayoutEntityType.ENDNOTE:
                childEntity = new RenderedEndnote();
                break;
            case LayoutEntityType.FOOTNOTE:
                childEntity = new RenderedFootnote();
                break;
            case LayoutEntityType.HEADER_FOOTER:
                childEntity = new RenderedHeaderFooter();
                break;
            case LayoutEntityType.LINE:
                childEntity = new RenderedLine();
                break;
            case LayoutEntityType.NOTE_SEPARATOR:
                childEntity = new RenderedNoteSeparator();
                break;
            case LayoutEntityType.PAGE:
                childEntity = new RenderedPage();
                break;
            case LayoutEntityType.ROW:
                childEntity = new RenderedRow();
                break;
            case LayoutEntityType.SPAN:
                childEntity = new RenderedSpan(it.getText());
                break;
            case LayoutEntityType.TEXT_BOX:
                childEntity = new RenderedTextBox();
                break;
            default:
                throw new IllegalStateException("Unknown layout type");
        }

        childEntity.mKind = it.getKind();
        childEntity.mPageIndex = it.getPageIndex();
        childEntity.mRectangle = it.getRectangleInternal();
        childEntity.mType = it.getType();
        childEntity.setLayoutObject(it.getCurrent());
        childEntity.mParent = this;

        return childEntity;
    }

    /// <summary>
    /// Returns a collection of child entities which match the specified type.
    /// </summary>
    /// <param name="type">Specifies the type of entities to select.</param>
    /// <param name="isDeep">True to select from all child entities recursively.
    /// False to select only among immediate children</param>
    public LayoutCollection<LayoutEntity> getChildEntities(/*LayoutEntityType*/int type, boolean isDeep)
    {
        ArrayList<LayoutEntity> childList = new ArrayList<LayoutEntity>();

        for (LayoutEntity entity : mChildEntities)
        {
            if ((entity.Type & type) == entity.Type)
                childList.add(entity);

            if (isDeep)
                msArrayList.addRange(childList, entity.getChildEntities(type, true));
        }

        return new LayoutCollection<LayoutEntity>(childList);
    }

    protected <T extends LayoutEntity> LayoutCollection<T> getChildNodes()
    {
        T obj = new T();
        ArrayList<T> childList = mChildEntities.Where(entity => entity.GetType() == obj.GetType()).<T>Cast().ToList();

        return new LayoutCollection<T>(childList);
    }

    protected String mKind;
    protected int mPageIndex;
    protected Node mParentNode;
    protected RectangleF mRectangle;
    protected /*LayoutEntityType*/int mType;
    protected LayoutEntity mParent;
    protected ArrayList<LayoutEntity> mChildEntities = new ArrayList<LayoutEntity>();
}

/// <summary>
/// Represents a generic collection of layout entity types.
/// </summary>
public final class LayoutCollection<T extends LayoutEntity> implements Iterable<T>
{
    /// <summary>
    /// Reserved for internal use.
    /// </summary>
    LayoutCollection(ArrayList<T> baseList)
    {
        mBaseList = baseList;
    }

    /// <summary>
    /// Provides a simple "foreach" style iteration over the collection of nodes. 
    /// </summary>
    //JAVA-deleted non-generic namesake:
    //Iterator iterator()

    /// <summary>
    /// Provides a simple "foreach" style iteration over the collection of nodes. 
    /// </summary>
    public Iterator<T> /*IEnumerable<T>.*/iterator()
    {
        return mBaseList.GetEnumerator();
    }

    /// <summary>
    /// Returns the first entity in the collection.
    /// </summary>
    public T First => mBaseList.Count > 0 ? mBaseList[0] : default;

    /// <summary>
    /// Returns the last entity in the collection.
    /// </summary>
    public T Last => mBaseList.Count > 0 ? mBaseList[mBaseList.Count - 1] : default;

    /// <summary>
    /// Retrieves the entity at the given index. 
    /// </summary>
    /// <remarks><para>The index is zero-based.</para>
    /// <para>If index is greater than or equal to the number of items in the list,
    /// this returns a null reference.</para></remarks>
     !!Autoporter error: Indexer DocsExamples.Complex_examples_and_helpers.LayoutCollection<T>.Item(int) hasn't both getter and setter! => index < mBaseList.Count ? mBaseList[index] : default;

    /// <summary>
    /// Gets the number of entities in the collection.
    /// </summary>
    public int Count => mBaseList.Count;

    private /*final*/ List<T> mBaseList;
}

/// <summary>
/// Represents an entity that contains lines and rows.
/// </summary>
public abstract class StoryLayoutEntity extends LayoutEntity
{private GetChildNodes<RenderedLine>getChildNodes();private GetChildNodes<RenderedRow>getChildNodes();
}

/// <summary>
/// Represents line of characters of text and inline objects.
/// </summary>
public class RenderedLine extends LayoutEntity
{private Environment.NewLineEnvironment;private ParentNodeParentNode;

    /// <summary>
    /// Provides access to the spans of the line.
    /// </summary>
    public LayoutCollection<RenderedSpan> Spans => private GetChildNodes<RenderedSpan>getChildNodes();
}

/// <summary>
/// Represents one or more characters in a line.
/// This include special characters like field start/end markers, bookmarks, shapes and comments.
/// </summary>
public class RenderedSpan extends LayoutEntity
{
    public RenderedSpan()
    {
    }

    RenderedSpan(String text)
    {
        // Assign empty text if the span text is null (this can happen with shape spans).
        mText = (text != null ? text : "");
    }private mKindmKind;

    /// <summary>
    /// Exports the contents of the entity into a string in plain text format.
    /// </summary>
    public /*override*/ String getText() { return mText; };

    private  String mText;private mParentNodemParentNode;
}

/// <summary>
/// Represents the header/footer content on a page.
/// </summary>
public class RenderedHeaderFooter extends StoryLayoutEntity
{
    /// <summary>
    /// Returns the type of the header or footer.
    /// </summary>
    public String Kind => private mKindmKind;
}

/// <summary>
/// Represents page of a document.
/// </summary>
public class RenderedPage extends LayoutEntity
{
    /// <summary>
    /// Provides access to the columns of the page.
    /// </summary>
    public LayoutCollection<RenderedColumn> Columns => private GetChildNodes<RenderedColumn>getChildNodes();private GetChildNodes<RenderedHeaderFooter>getChildNodes();

    /// <summary>
    /// Provides access to the comments of the page.
    /// </summary>
    public LayoutCollection<RenderedComment> Comments => private GetChildNodes<RenderedComment>getChildNodes();

    /// <summary>
    /// Returns the section that corresponds to the layout entity.  
    /// </summary>
    public Section Section => (Section) private ParentNodeParentNode;private Columns.First.GetChildEntitiescolumns(LayoutEntityType.Line, true)private First.ParentNode.GetAncestorfirst(NodeType.Section);
}

/// <summary>
/// Represents a table row.
/// </summary>
public class RenderedRow extends LayoutEntity
{private GetChildNodes<RenderedCell>getChildNodes();

    /// <summary>
    /// Returns the row that corresponds to the layout entity.  
    /// </summary>
    /// <remarks>This property may return null for some rows such as those inside the header or footer.</remarks>
    public Row Row => (Row) private ParentNodeParentNode;private ParentTableParentTable;

    /// <summary>
    /// Returns the node that corresponds to this layout entity.  
    /// </summary>
    /// <remarks>This property may return null for nodes that are inside the header or footer.</remarks>
    public /*override*/ Node getParentNode()
    {
        Paragraph para = Cells.First.Lines.First?.Paragraph;
        return para?.GetAncestor(NodeType.Row);
    }
}

/// <summary>
/// Represents a column of text on a page.
/// </summary>
public class RenderedColumn extends StoryLayoutEntity
{private GetChildNodes<RenderedFootnote>getChildNodes();

    /// <summary>
    /// Provides access to the endnotes of the page.
    /// </summary>
    public LayoutCollection<RenderedEndnote> Endnotes => private GetChildNodes<RenderedEndnote>getChildNodes();

    /// <summary>
    /// Provides access to the note separators of the page.
    /// </summary>
    public LayoutCollection<RenderedNoteSeparator> NoteSeparators => private GetChildNodes<RenderedNoteSeparator>getChildNodes();private ParentNodeParentNode;

    /// <summary>
    /// Returns the node that corresponds to this layout entity.  
    /// </summary>
    public /*override*/ Node ParentNode => 
        private GetChildEntitiesgetChildEntities(LayoutEntityType.Line, true).private First.ParentNode.GetAncestorfirst(NodeType.Body);
}

/// <summary>
/// Represents a table cell.
/// </summary>
public class RenderedCell extends StoryLayoutEntity
{
    /// <summary>
    /// Returns the cell that corresponds to the layout entity.  
    /// </summary>
    /// <remarks>This property may return null for some cells such as those inside the header or footer.</remarks>
    public Cell Cell => (Cell) private ParentNodeParentNode;private GetAncestorgetAncestor(NodeType.Cell);
}

/// <summary>
/// Represents placeholder for footnote content.
/// </summary>
public class RenderedFootnote extends StoryLayoutEntity
{private ParentNodeParentNode;

    /// <summary>
    /// Returns the node that corresponds to this layout entity.  
    /// </summary>
    public /*override*/ Node ParentNode => 
        private GetChildEntitiesgetChildEntities(LayoutEntityType.Line, true).private First.ParentNode.GetAncestorfirst(NodeType.Footnote);
}

/// <summary>
/// Represents placeholder for endnote content.
/// </summary>
public class RenderedEndnote extends StoryLayoutEntity
{
    /// <summary>
    /// Returns the endnote that corresponds to the layout entity.  
    /// </summary>
    public Footnote Endnote => (Footnote) private ParentNodeParentNode;private GetChildEntitiesgetChildEntities(LayoutEntityType.Line, true).private First.ParentNode.GetAncestorfirst(NodeType.Footnote);
}

/// <summary>
/// Represents text area inside of a shape.
/// </summary>
public class RenderedTextBox extends StoryLayoutEntity
{
    /// <summary>
    /// Returns the Shape or DrawingML that corresponds to the layout entity.  
    /// </summary>
    /// <remarks>This property may return null for some Shapes or DrawingML such as those inside the header or footer.</remarks>
    public /*override*/ Node getParentNode()
    {
        LayoutCollection<LayoutEntity> lines = getChildEntities(LayoutEntityType.LINE, true);
        Node shape = lines.First.ParentNode.GetAncestor(NodeType.SHAPE);

        return (shape != null ? shape : lines.First.ParentNode.GetAncestor(NodeType.SHAPE));
    }
}

/// <summary>
/// Represents placeholder for comment content.
/// </summary>
public class RenderedComment extends StoryLayoutEntity
{private ParentNodeParentNode;

    /// <summary>
    /// Returns the node that corresponds to this layout entity.  
    /// </summary>
    public /*override*/ Node ParentNode => 
        private GetChildEntitiesgetChildEntities(LayoutEntityType.Line, true).private First.ParentNode.GetAncestorfirst(NodeType.Comment);
}

/// <summary>
/// Represents footnote/endnote separator.
/// </summary>
public class RenderedNoteSeparator extends StoryLayoutEntity
{
    /// <summary>
    /// Returns the footnote/endnote that corresponds to the layout entity.  
    /// </summary>
    public Footnote Footnote => (Footnote) private ParentNodeParentNode;private GetChildEntitiesgetChildEntities(LayoutEntityType.Line, true).private First.ParentNode.GetAncestorfirst(NodeType.Footnote);
}
