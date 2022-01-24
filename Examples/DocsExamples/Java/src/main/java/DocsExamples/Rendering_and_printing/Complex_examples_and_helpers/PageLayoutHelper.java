package DocsExamples.Rendering_and_printing.Complex_examples_and_helpers;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

import java.awt.geom.Rectangle2D;
import java.text.MessageFormat;
import java.util.List;
import java.util.*;

@Test
public class PageLayoutHelper extends DocsExamplesBase
{
    @Test
    public void wrapperToAccessLayoutEntities() throws Exception
    {
        // This sample introduces the RenderedDocument class and other related classes which provide an API wrapper for 
        // the LayoutEnumerator. This allows you to access the layout entities of a document using a DOM style API.
        Document doc = new Document(getMyDir() + "Document layout.docx");

        RenderedDocument layoutDoc = new RenderedDocument(doc);

        // Get access to the line of the first page and print to the console.
        RenderedLine line = layoutDoc.getPages().getItem(0).getColumns().getItem(0).getLines().getItem(2);
        System.out.println("Line: " + line.getText());

        // With a rendered line, the original paragraph in the document object model can be returned.
        Paragraph para = line.getParagraph();
        System.out.println("Paragraph text: " + para.getRange().getText());

        // Retrieve all the text that appears on the first page in plain text format (including headers and footers).
        String pageText = layoutDoc.getPages().getItem(0).getText();
        System.out.println();

        // Loop through each page in the document and print how many lines appear on each page.
        for (RenderedPage page : layoutDoc.getPages())
        {
            LayoutCollection<LayoutEntity> lines = page.getChildEntities(LayoutEntityType.LINE, true);
            System.out.println(MessageFormat.format("Page {0} has {1} lines.", page.getPageIndex(), lines.getCount()));
        }

        // This method provides a reverse lookup of layout entities for any given node
        // (except runs and nodes in the header and footer).
        System.out.println();
        System.out.println("The lines of the second paragraph:");
        for (LayoutEntity layoutEntity : layoutDoc.getLayoutEntitiesOfNode(
            doc.getFirstSection().getBody().getParagraphs().get(1)))
        {
            RenderedLine paragraphLine = (RenderedLine) layoutEntity;
            System.out.println("\"{paragraphLine.Text.Trim()}\"");
            System.out.println(paragraphLine.getRectangle().toString());
            System.out.println();
        }
    }
}

/// <summary>
/// Provides an API wrapper for the LayoutEnumerator class to access the page layout
/// of a document presented in an object model like the design.
/// </summary>
class RenderedDocument extends LayoutEntity
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
    public final LayoutCollection<RenderedPage> getPages() {
        return getChildNodes(new RenderedPage());
    }

    /// <summary>
    /// Returns all the layout entities of the specified node.
    /// </summary>
    /// <remarks>Note that this method does not work with Run nodes or nodes in the header or footer.</remarks>
    LayoutCollection<LayoutEntity> getLayoutEntitiesOfNode(Node node)
    {
        if (!mLayoutCollector.getDocument().equals(node.getDocument()))
            throw new IllegalArgumentException("Node does not belong to the same document which was rendered.");

        if (node.getNodeType() == NodeType.DOCUMENT)
            return new LayoutCollection<>(mChildEntities);

        ArrayList<LayoutEntity> entities = new ArrayList<LayoutEntity>();

        // Retrieve all entities from the layout document (inversion of LayoutEntityType.None).
        for (LayoutEntity entity : getChildEntities(~LayoutEntityType.NONE, true))
        {
            if (entity.getParentNode() == node)
                entities.add(entity);

            // There is no table entity in rendered output, so manually check if rows belong to a table node.
            if (entity.getType() == LayoutEntityType.ROW)
            {
                RenderedRow row = (RenderedRow) entity;
                if (row.getTable() == node)
                    entities.add(entity);
            }
        }

        return new LayoutCollection<>(entities);
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

                current = current.getParent();
            }
        } while (mEnumerator.moveNext());
    }

    private void collectLinesAndAddToMarkers()
    {
        collectLinesOfMarkersCore(LayoutEntityType.COLUMN);
        collectLinesOfMarkersCore(LayoutEntityType.COMMENT);
    }

    private void collectLinesOfMarkersCore(int type)
    {
        ArrayList<RenderedLine> collectedLines = new ArrayList<>();

        for (RenderedPage page : getPages())
        {
            for (LayoutEntity story : page.getChildEntities(type, false))
            {
                for (LayoutEntity le : story.getChildEntities(LayoutEntityType.LINE, true))
                {
                    RenderedLine line = (RenderedLine) le;
                    collectedLines.add(line);
                    for (RenderedSpan span : line.getSpans())
                    {
                        if (mLayoutToNodeLookup.containsKey(span.getLayoutObject()))
                        {
                            if ("PARAGRAPH".equals(span.getKind()) || "ROW".equals(span.getKind()) || "CELL".equals(span.getKind()) ||
                                "SECTION".equals(span.getKind()))
                            {
                                Node node = mLayoutToNodeLookup.get(span.getLayoutObject());

                                if (node.getNodeType() == NodeType.ROW)
                                    node = ((Row) node).getLastCell().getLastParagraph();

                                for (RenderedLine collectedLine : collectedLines)
                                    collectedLine.setParentNode(node);

                                collectedLines = new ArrayList<>();
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
                mLayoutToNodeLookup.put(entity, node);
        }
    }

    private LayoutCollector mLayoutCollector;
    private LayoutEnumerator mEnumerator;
    private HashMap<Object, Node> mLayoutToNodeLookup = new HashMap<>();
}

/// <summary>
/// Provides the base class for rendered elements of a document.
/// </summary>
abstract class LayoutEntity
{
    public final int getPageIndex() {
        return mPageIndex;
    }

    /// <summary>
    /// Returns bounding rectangle of the entity relative to the page top left corner (in points).
    /// </summary>
    public final Rectangle2D getRectangle() {
        return mRectangle;
    }

    public final int getType() { return mType; }

    /// <summary>
    /// Exports the contents of the entity into a string in plain text format.
    /// </summary>
    public String getText()
    {
        StringBuilder builder = new StringBuilder();
        for (LayoutEntity entity : mChildEntities)
        {
            builder.append(entity.getText());
        }

        return builder.toString();
    }

    public final LayoutEntity getParent() {
        return mParent;
    }

    /// <summary>
    /// Returns the node that corresponds to this layout entity.  
    /// </summary>
    /// <remarks>This property may return null for spans that originate
    /// from Run nodes or nodes inside the header or footer.</remarks>
    public Node getParentNode() {
        return mParentNode;
    }

    /// <summary>
    /// Internal method separate from ParentNode property to make code autoportable to VB.NET.
    /// </summary>
    public void setParentNode(Node value)
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
        childEntity.mRectangle = it.getRectangle();
        childEntity.mType = it.getType();
        childEntity.setLayoutObject(it.getCurrent());
        childEntity.mParent = this;

        return childEntity;
    }

    public static <E> Collection<E> makeCollection(Iterable<E> iter) {
        Collection<E> list = new java.util.ArrayList<E>();
        for (E item : iter) {
            list.add(item);
        }
        return list;
    }

    /// <summary>
    /// Returns a collection of child entities which match the specified type.
    /// </summary>
    /// <param name="type">Specifies the type of entities to select.</param>
    /// <param name="isDeep">True to select from all child entities recursively.
    /// False to select only among immediate children</param>
    public LayoutCollection<LayoutEntity> getChildEntities(int type, boolean isDeep)
    {
        ArrayList<LayoutEntity> childList = new ArrayList<>();

        for (LayoutEntity entity : mChildEntities)
        {
            if ((entity.getType() & type) == entity.getType())
                childList.add(entity);

            if (isDeep)
                childList.addAll(makeCollection(entity.getChildEntities(type, true)));
        }

        return new LayoutCollection<>(childList);
    }

    protected <T extends LayoutEntity> LayoutCollection<T> getChildNodes(T t)
    {
        T obj = t;
        ArrayList<T> childList = new ArrayList<>();

        for (LayoutEntity entity : mChildEntities) {
            if (entity.getClass() == obj.getClass()) {
                System.out.println(obj.getClass());
                childList.add((T) entity);
            }
        }

        return new LayoutCollection<>(childList);
    }

    protected String mKind;
    protected int mPageIndex;
    protected Node mParentNode;
    protected Rectangle2D mRectangle;
    protected int mType;
    protected LayoutEntity mParent;
    protected ArrayList<LayoutEntity> mChildEntities = new ArrayList<>();
}

/// <summary>
/// Represents a generic collection of layout entity types.
/// </summary>
final class LayoutCollection<T extends LayoutEntity> implements Iterable<T>
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
    public final java.util.Iterator getEnumerator() {
        return mBaseList.iterator();
    }

    /// <summary>
    /// Provides a simple "foreach" style iteration over the collection of nodes. 
    /// </summary>
    public Iterator<T> iterator()
    {
        return mBaseList.iterator();
    }

    /// <summary>
    /// Returns the first entity in the collection.
    /// </summary>
    public final T getFirst() {
        if (mBaseList.size() > 0) {
            return mBaseList.get(0);
        } else {
            return null;
        }
    }

    /// <summary>
    /// Returns the last entity in the collection.
    /// </summary>
    public final T getLast() {
        if (mBaseList.size() > 0) {
            return mBaseList.get(mBaseList.size() - 1);
        } else {
            return null;
        }
    }

    /// <summary>
    /// Retrieves the entity at the given index. 
    /// </summary>
    /// <remarks><para>The index is zero-based.</para>
    /// <para>If index is greater than or equal to the number of items in the list,
    /// this returns a null reference.</para></remarks>
    public final T getItem(int index) {
        return mBaseList.get(index);
    }

    /// <summary>
    /// Gets the number of entities in the collection.
    /// </summary>
    public final int getCount() {
        return mBaseList.size();
    }

    private List<T> mBaseList;
}

/// <summary>
/// Represents an entity that contains lines and rows.
/// </summary>
abstract class StoryLayoutEntity extends LayoutEntity
{
    /// <summary>
    /// Provides access to the lines of a story.
    /// </summary>
    public final LayoutCollection<RenderedLine> getLines() {
        return getChildNodes(new RenderedLine());
    }

    /// <summary>
    /// Provides access to the row entities of a table.
    /// </summary>
    public final LayoutCollection<RenderedRow> getRows() {
        return getChildNodes(new RenderedRow());
    }
}

/// <summary>
/// Represents line of characters of text and inline objects.
/// </summary>
class RenderedLine extends LayoutEntity
{
    @Override
    public String getText() {
        return super.getText() + "\n";
    }

    public final Paragraph getParagraph() {
        return (Paragraph) getParentNode();
    }

    /// <summary>
    /// Provides access to the spans of the line.
    /// </summary>
    public final LayoutCollection<RenderedSpan> getSpans() {
        return getChildNodes(new RenderedSpan());
    }
}

/// <summary>
/// Represents one or more characters in a line.
/// This include special characters like field start/end markers, bookmarks, shapes and comments.
/// </summary>
class RenderedSpan extends LayoutEntity
{
    public RenderedSpan()
    {
    }

    RenderedSpan(String text)
    {
        // Assign empty text if the span text is null (this can happen with shape spans).
        mText = (text != null ? text : "");
    }

    public final String getKind() {
        return mKind;
    }

    /// <summary>
    /// Exports the contents of the entity into a string in plain text format.
    /// </summary>
    @Override
    public String getText() {
        return mText;
    }

    @Override
    public Node getParentNode() {
        return mParentNode;
    }

    private String mText;
}

/// <summary>
/// Represents the header/footer content on a page.
/// </summary>
class RenderedHeaderFooter extends StoryLayoutEntity
{
    /// <summary>
    /// Returns the type of the header or footer.
    /// </summary>
    public final String getKind() {
        return mKind;
    }
}

/// <summary>
/// Represents page of a document.
/// </summary>
class RenderedPage extends LayoutEntity
{
    /// <summary>
    /// Provides access to the columns of the page.
    /// </summary>
    public final LayoutCollection<RenderedColumn> getColumns() {
        return getChildNodes(new RenderedColumn());
    }

    public final LayoutCollection<RenderedHeaderFooter> getHeaderFooters() {
        return getChildNodes(new RenderedHeaderFooter());
    }

    /// <summary>
    /// Provides access to the comments of the page.
    /// </summary>
    public final LayoutCollection<RenderedComment> getComments() {
        return getChildNodes(new RenderedComment());
    }

    /// <summary>
    /// Returns the section that corresponds to the layout entity.  
    /// </summary>
    public final Section getSection() {
        return (Section) getParentNode();
    }

    @Override
    public Node getParentNode() {
        return getColumns().getFirst().getLines().getFirst().getParagraph().getParentSection();
    }
}

/// <summary>
/// Represents a table row.
/// </summary>
class RenderedRow extends LayoutEntity
{
    public final LayoutCollection<RenderedCell> getCells() {
        return getChildNodes(new RenderedCell());
    }

    /// <summary>
    /// Returns the row that corresponds to the layout entity.  
    /// </summary>
    /// <remarks>This property may return null for some rows such as those inside the header or footer.</remarks>
    public final Row getRow() {
        return (Row) getParentNode();
    }

    public final Table getTable() {
        return getRow().getParentTable();
    }

    /// <summary>
    /// Returns the node that corresponds to this layout entity.  
    /// </summary>
    /// <remarks>This property may return null for nodes that are inside the header or footer.</remarks>
    @Override
    public Node getParentNode() {
        return getCells().getFirst().getLines().getFirst().getParagraph().getAncestor(NodeType.ROW);
    }
}

/// <summary>
/// Represents a column of text on a page.
/// </summary>
class RenderedColumn extends StoryLayoutEntity
{
    public final LayoutCollection<RenderedFootnote> getFootnotes() {
        return getChildNodes(new RenderedFootnote());
    }

    /// <summary>
    /// Provides access to the endnotes of the page.
    /// </summary>
    public final LayoutCollection<RenderedEndnote> getEndnotes() {
        return getChildNodes(new RenderedEndnote());
    }

    /// <summary>
    /// Provides access to the note separators of the page.
    /// </summary>
    public final LayoutCollection<RenderedNoteSeparator> getNoteSeparators() {
        return getChildNodes(new RenderedNoteSeparator());
    }

    public final Body getBody() {
        return (Body) getParentNode();
    }

    /// <summary>
    /// Returns the node that corresponds to this layout entity.  
    /// </summary>
    @Override
    public Node getParentNode() {
        return getLines().getFirst().getParagraph().getParentSection().getBody();
    }
}

/// <summary>
/// Represents a table cell.
/// </summary>
class RenderedCell extends StoryLayoutEntity
{
    public final Cell getCell() {
        return (Cell) getParentNode();
    }

    /// <summary>
    /// Returns the cell that corresponds to the layout entity.  
    /// </summary>
    /// <remarks>This property may return null for some cells such as those inside the header or footer.</remarks>
    @Override
    public Node getParentNode() {
        return getLines().getFirst().getParagraph().getAncestor(NodeType.CELL);
    }
}

/// <summary>
/// Represents placeholder for footnote content.
/// </summary>
class RenderedFootnote extends StoryLayoutEntity
{
    public final Footnote getFootnote() {
        return (Footnote) getParentNode();
    }

    /// <summary>
    /// Returns the node that corresponds to this layout entity.  
    /// </summary>
    @Override
    public Node getParentNode() {
        return getLines().getFirst().getParagraph().getAncestor(NodeType.FOOTNOTE);
    }
}

/// <summary>
/// Represents placeholder for endnote content.
/// </summary>
class RenderedEndnote extends StoryLayoutEntity
{
    /// <summary>
    /// Returns the endnote that corresponds to the layout entity.  
    /// </summary>
    public final Footnote getEndnote() {
        return (Footnote) getParentNode();
    }

    @Override
    public Node getParentNode() {
        return getLines().getFirst().getParagraph().getAncestor(NodeType.FOOTNOTE);
    }
}

/// <summary>
/// Represents text area inside of a shape.
/// </summary>
class RenderedTextBox extends StoryLayoutEntity
{
    /// <summary>
    /// Returns the Shape or DrawingML that corresponds to the layout entity.  
    /// </summary>
    /// <remarks>This property may return null for some Shapes or DrawingML such as those inside the header or footer.</remarks>
    @Override
    public Node getParentNode() {
        Node shape = getLines().getFirst().getParagraph().getAncestor(NodeType.SHAPE);

        if (shape != null) {
            return shape;
        } else
            return null;
    }
}

/// <summary>
/// Represents placeholder for comment content.
/// </summary>
class RenderedComment extends StoryLayoutEntity
{
    public final Comment getComment() {
        return (Comment) getParentNode();
    }

    /// <summary>
    /// Returns the node that corresponds to this layout entity.  
    /// </summary>
    @Override
    public Node getParentNode() {
        return getLines().getFirst().getParagraph().getAncestor(NodeType.COMMENT);
    }
}

/// <summary>
/// Represents footnote/endnote separator.
/// </summary>
class RenderedNoteSeparator extends StoryLayoutEntity
{
    /// <summary>
    /// Returns the footnote/endnote that corresponds to the layout entity.  
    /// </summary>
    public final Footnote getFootnote() {
        return (Footnote) getParentNode();
    }

    @Override
    public Node getParentNode() {
        return getLines().getFirst().getParagraph().getAncestor(NodeType.FOOTNOTE);
    }
}
