package com.aspose.words.examples.rendering_printing;

import com.aspose.words.*;

import java.awt.geom.Rectangle2D;
import java.util.Collection;

/**
 * Provides the base class for rendered elements of a document.
 */
public class LayoutEntity {
    protected LayoutEntity() {
    }

    /**
     * Gets the 1-based index of a page which contains the rendered entity.
     */
    public final int getPageIndex() {
        return mPageIndex;
    }

    /**
     * Returns bounding rectangle of the entity relative to the page top left corner (in points).
     */
    public final Rectangle2D getRectangle() {
        return mRectangle;
    }

    /**
     * Gets the type of this layout entity.
     */
    public final int getType() //LayoutEntityType
    {
        return mType;
    }

    /**
     * Exports the contents of the entity into a string in plain text format.
     */
    public String getText() {
        StringBuilder builder = new StringBuilder();
        for (LayoutEntity entity : mChildEntities) {
            builder.append(entity.getText());
        }

        return builder.toString();
    }

    /**
     * Gets the immediate parent of this entity.
     */
    public final LayoutEntity getParent() {
        return mParent;
    }

    /**
     * Returns the node that corresponds to this layout entity.
     * <p>
     * This property may return null for spans that originate from Run nodes or nodes that are inside the header or footer.
     */
    public Node getParentNode() {
        return mParentNode;
    }

    public void setParentNode(Node value) {
        //System.out.println(value);
        mParentNode = value;
    }

    /**
     * Reserved for internal use.
     */
    private Object privateLayoutObject;

    public final Object getLayoutObject() {
        return privateLayoutObject;
    }

    public final void setLayoutObject(Object value) {
        privateLayoutObject = value;
    }

    /**
     * Reserved for internal use.
     *
     * @throws Exception
     */
    public final LayoutEntity AddChildEntity(LayoutEnumerator it) throws Exception {
        LayoutEntity child = CreateLayoutEntityFromType(it);
        mChildEntities.add(child);

        return child;
    }


    private LayoutEntity CreateLayoutEntityFromType(LayoutEnumerator it) throws Exception {
        LayoutEntity childEntity;
        switch (it.getType()) {
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
                throw new UnsupportedOperationException("Unknown layout type");
        }

        childEntity.mKind = it.getKind();
        childEntity.mPageIndex = it.getPageIndex();
        childEntity.mRectangle = it.getRectangle();
        childEntity.mType = it.getType();
        childEntity.setLayoutObject(it.getCurrent());
        childEntity.mParent = this;

        return childEntity;
    }

    /**
     * Returns a collection of child entities which match the specified type.
     */

    public static <E> Collection<E> makeCollection(Iterable<E> iter) {
        Collection<E> list = new java.util.ArrayList<E>();
        for (E item : iter) {
            list.add(item);
        }
        return list;
    }

    @SuppressWarnings("unchecked")
    public final LayoutCollection<LayoutEntity> GetChildEntities(int type, boolean isDeep) {
        //System.out.println("GetChildEntities");System.out.println(type);
        java.util.ArrayList<LayoutEntity> childList = new java.util.ArrayList<LayoutEntity>();
        //Iterable<? extends LayoutEntity> childList = new java.util.ArrayList<LayoutEntity>();

        for (LayoutEntity entity : mChildEntities) {
            if ((entity.getType() & type) == entity.getType()) {
                childList.add(entity);
            }

            if (isDeep) {

                //System.out.println(entity.GetChildEntities(type, true));
                Iterable<LayoutEntity> collection = (Iterable<LayoutEntity>) entity.GetChildEntities(type, true);
                childList.addAll(makeCollection(collection));
                //childList.addAll(com.google.common.collect.Lists.newArrayList(collection));
            }
        }

        return new LayoutCollection<LayoutEntity>(childList);
    }


    @SuppressWarnings("unchecked")
    public <T extends LayoutEntity> LayoutCollection<T> GetChildNodes(T t) //<T> T GetChildNodes(String name) //<T extends LayoutEntity> LayoutCollection<T> GetChildNodes(T t)
    {
        T obj = t;

        java.util.ArrayList<T> childList = new java.util.ArrayList<T>();

        for (LayoutEntity entity : mChildEntities) {
            //System.out.println(entity);
			/*
			switch (name) {
			 case "RenderedLine":
				 childList.add((RenderedLine)entity);
                     break;
			}
			*/

            if (entity.getClass() == obj.getClass()) {
                System.out.println(obj.getClass());
                childList.add((T) entity);
            }
        }

        return (LayoutCollection<T>) new LayoutCollection(childList);
    }


    protected String mKind;
    protected int mPageIndex;
    protected Node mParentNode;
    protected Rectangle2D mRectangle;
    protected int mType; //LayoutEntityType
    protected LayoutEntity mParent;
    protected java.util.ArrayList<LayoutEntity> mChildEntities = new java.util.ArrayList<LayoutEntity>();
}


//////////////////////////////////////

/**
 * Represents a generic collection of layout entity types.
 */
class LayoutCollection<T> implements Iterable<T> // extends LayoutEntity>
{
    /**
     * Reserved for internal use.
     */
    public LayoutCollection(java.util.ArrayList<T> baseList) {
        mBaseList = baseList;
    }

    /**
     * Provides a simple "foreach" style iteration over the collection of nodes.
     */
    public final java.util.Iterator GetEnumerator() {
        return mBaseList.iterator();
    }

    /**
     * Provides a simple "foreach" style iteration over the collection of nodes.
     */
    public final java.util.Iterator<T> iterator() {
        return mBaseList.iterator();
    }

    /**
     * Returns the first entity in the collection.
     */
    public final T getFirst() {
        if (mBaseList.size() > 0) {
            return mBaseList.get(0);
        } else {
            return null;
        }
    }

    /**
     * Returns the last entity in the collection.
     */
    public final T getLast() {
        if (mBaseList.size() > 0) {
            return mBaseList.get(mBaseList.size() - 1);
        } else {
            return null;
        }
    }

    /**
     * Retrieves the entity at the given index.
     * <p>
     * <p>The index is zero-based.</p>
     * <p>If index is greater than or equal to the number of items in the list, this returns a null reference.</p>
     */
    public final T getItem(int index) {
        return mBaseList.get(index);
    }

    /**
     * Gets the number of entities in the collection.
     */
    public final int getCount() {
        return mBaseList.size();
    }

    private java.util.ArrayList<T> mBaseList;
}

/**
 * Represents an entity that contains lines and rows.
 */
class StoryLayoutEntity extends LayoutEntity {
    /**
     * Provides access to the lines of a story.
     */
    public final LayoutCollection<RenderedLine> getLines() {
        return GetChildNodes(new RenderedLine());
    }

    /**
     * Provides access to the row entities of a table.
     */
    public final LayoutCollection<RenderedRow> getRows() {
        return GetChildNodes(new RenderedRow());
    }
}

/**
 * Represents line of characters of text and inline objects.
 */
class RenderedLine extends LayoutEntity {
    /**
     * Exports the contents of the entity into a string in plain text format.
     */
    @Override
    public String getText() {
        return super.getText() + "\n";
    }

    /**
     * Returns the paragraph that corresponds to the layout entity.
     * <p>
     * This property may return null for some lines such as those inside the header or footer.
     */
    public final Paragraph getParagraph() {
        return (Paragraph) getParentNode();
    }

    /**
     * Provides access to the spans of the line.
     */
    public final LayoutCollection<RenderedSpan> getSpans() {
        return GetChildNodes(new RenderedSpan());
    }
}

/**
 * Represents one or more characters in a line.
 * This include special characters like field start/end markers, bookmarks and comments.
 */
class RenderedSpan extends LayoutEntity {
    public RenderedSpan() {
    }

    public RenderedSpan(String text) {
        mText = text;
    }

    /**
     * Gets kind of the span. This cannot be null.
     * <p>
     * This is a more specific type of the current entity, e.g. bookmark span has Span type and
     * may have either a BOOKMARKSTART or BOOKMARKEND kind.
     */
    public final String getKind() {
        return mKind;
    }

    /**
     * Exports the contents of the entity into a string in plain text format.
     */
    @Override
    public String getText() {
        return mText;
    }

    /**
     * Returns the node that corresponds to this layout entity.
     * <p>
     * This property returns null for spans that originate from Run nodes or nodes that are inside the header or footer.
     */
    @Override
    public Node getParentNode() {
        return mParentNode;
    }

    private String mText;
}

/**
 * Represents the header/footer content on a page.
 */
class RenderedHeaderFooter extends StoryLayoutEntity {
    /**
     * Returns the type of the header or footer.
     */
    public final String getKind() {
        return mKind;
    }
}

/**
 * Represents page of a document.
 */
class RenderedPage extends LayoutEntity {
    /**
     * Provides access to the columns of the page.
     */
    public final LayoutCollection<RenderedColumn> getColumns() {
        return GetChildNodes(new RenderedColumn());
    }

    /**
     * Provides access to the header and footers of the page.
     */
    public final LayoutCollection<RenderedHeaderFooter> getHeaderFooters() {
        return GetChildNodes(new RenderedHeaderFooter());
    }

    /**
     * Provides access to the comments of the page.
     */
    public final LayoutCollection<RenderedComment> getComments() {
        return GetChildNodes(new RenderedComment());
    }

    /**
     * Returns the section that corresponds to the layout entity.
     */
    public final Section getSection() {
        return (Section) getParentNode();
    }

    /**
     * Returns the node that corresponds to this layout entity.
     */
    @Override
    public Node getParentNode() {
        return getColumns().getFirst().getLines().getFirst().getParagraph().getParentSection();
    }
}

/**
 * Represents a table row.
 */
class RenderedRow extends LayoutEntity {
    /**
     * Provides access to the cells of the row.
     */
    public final LayoutCollection<RenderedCell> getCells() {
        return GetChildNodes(new RenderedCell());
    }

    /**
     * Returns the row that corresponds to the layout entity.
     * <p>
     * This property may return null for some rows such as those inside the header or footer.
     */
    public final Row getRow() {
        return (Row) getParentNode();
    }

    /**
     * Returns the table that corresponds to the layout entity.
     * <p>
     * This property may return null for some tables such as those inside the header or footer.
     */
    public final Table getTable() {
        return getRow().getParentTable();
    }

    /**
     * Returns the node that corresponds to this layout entity.
     * <p>
     * This property may return null for nodes that are inside the header or footer.
     */
    @Override
    public Node getParentNode() {
        return getCells().getFirst().getLines().getFirst().getParagraph().getAncestor(NodeType.ROW);
    }
}

/**
 * Represents a column of text on a page.
 */
class RenderedColumn extends StoryLayoutEntity {
    /**
     * Provides access to the footnotes of the page.
     */
    public final LayoutCollection<RenderedFootnote> getFootnotes() {
        return GetChildNodes(new RenderedFootnote());
    }

    /**
     * Provides access to the endnotes of the page.
     */
    public final LayoutCollection<RenderedEndnote> getEndnotes() {
        return GetChildNodes(new RenderedEndnote());
    }

    /**
     * Provides access to the note separators of the page.
     */
    public final LayoutCollection<RenderedNoteSeparator> getNoteSeparators() {
        return GetChildNodes(new RenderedNoteSeparator());
    }

    /**
     * Returns the body that corresponds to the layout entity.
     */
    public final Body getBody() {
        return (Body) getParentNode();
    }

    /**
     * Returns the node that corresponds to this layout entity.
     */
    @Override
    public Node getParentNode() {
        return getLines().getFirst().getParagraph().getParentSection().getBody();
    }
}

/**
 * Represents a table cell.
 */
class RenderedCell extends StoryLayoutEntity {
    /**
     * Returns the cell that corresponds to the layout entity.
     * <p>
     * This property may return null for some cells such as those inside the header or footer.
     */
    public final Cell getCell() {
        return (Cell) getParentNode();
    }

    /**
     * Returns the node that corresponds to this layout entity.
     * <p>
     * This property may return null for nodes that are inside the header or footer.
     */
    @Override
    public Node getParentNode() {
        return getLines().getFirst().getParagraph().getAncestor(NodeType.CELL);
    }
}

/**
 * Represents placeholder for footnote content.
 */
class RenderedFootnote extends StoryLayoutEntity {
    /**
     * Returns the footnote that corresponds to the layout entity.
     */
    public final Footnote getFootnote() {
        return (Footnote) getParentNode();
    }

    /**
     * Returns the node that corresponds to this layout entity.
     */
    @Override
    public Node getParentNode() {
        return getLines().getFirst().getParagraph().getAncestor(NodeType.FOOTNOTE);
    }
}

/**
 * Represents placeholder for endnote content.
 */
class RenderedEndnote extends StoryLayoutEntity {
    /**
     * Returns the endnote that corresponds to the layout entity.
     */
    public final Footnote getEndnote() {
        return (Footnote) getParentNode();
    }

    /**
     * Returns the node that corresponds to this layout entity.
     */
    @Override
    public Node getParentNode() {
        return getLines().getFirst().getParagraph().getAncestor(NodeType.FOOTNOTE);
    }
}

/**
 * Represents text area inside of a shape.
 */
class RenderedTextBox extends StoryLayoutEntity {
    /**
     * Returns the Shape or DrawingML that corresponds to the layout entity.
     * <p>
     * This property may return null for some Shapes or DrawingML such as those inside the header or footer.
     */
    @Override
    public Node getParentNode() {
        Node shape = getLines().getFirst().getParagraph().getAncestor(NodeType.SHAPE);

        if (shape != null) {
            return shape;
        } else
            return null;
    }
}

/**
 * Represents placeholder for comment content.
 */
class RenderedComment extends StoryLayoutEntity {
    /**
     * Returns the comment that corresponds to the layout entity.
     */
    public final Comment getComment() {
        return (Comment) getParentNode();
    }

    /**
     * Returns the node that corresponds to this layout entity.
     */
    @Override
    public Node getParentNode() {
        return getLines().getFirst().getParagraph().getAncestor(NodeType.COMMENT);
    }
}

/**
 * Represents footnote/endnote separator.
 */
class RenderedNoteSeparator extends StoryLayoutEntity {
    /**
     * Returns the footnote/endnote that corresponds to the layout entity.
     */
    public final Footnote getFootnote() {
        return (Footnote) getParentNode();
    }

    /**
     * Returns the node that corresponds to this layout entity.
     */
    @Override
    public Node getParentNode() {
        return getLines().getFirst().getParagraph().getAncestor(NodeType.FOOTNOTE);
    }
}