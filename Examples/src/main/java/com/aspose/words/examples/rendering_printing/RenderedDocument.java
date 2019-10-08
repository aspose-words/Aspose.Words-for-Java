package com.aspose.words.examples.rendering_printing;

import com.aspose.words.*;

public class RenderedDocument extends LayoutEntity {
    /**
     * Creates a new instance from the supplied Aspose.Words.Document class.
     *
     * @param document A document whose page layout model to enumerate.
     *                 <p>If page layout model of the document hasn't been built the enumerator calls <see cref="Document.UpdatePageLayout"/> to build it.</p>
     *                 <p>Whenever document is updated and new page layout model is created, a new enumerator must be used to access it.</p>
     * @throws Exception
     */
    public RenderedDocument(Document doc) throws Exception {
        mLayoutCollector = new LayoutCollector(doc);
        mEnumerator = new LayoutEnumerator(doc);
        ProcessLayoutElements(this);
        CollectLinesAndAddToMarkers();
        LinkLayoutMarkersToNodes(doc);
    }

    /**
     * Provides access to the pages of a document.
     */
    public final LayoutCollection<RenderedPage> getPages() {
        return GetChildNodes(new RenderedPage());
    }

    /**
     * Returns all the layout entities of the specified node.
     * <p>
     * Note that this method does not work with Run nodes or nodes in the header or footer.
     */
    public final LayoutCollection<LayoutEntity> GetLayoutEntitiesOfNode(Node node) {
        if (mLayoutCollector.getDocument() != node.getDocument()) {
            throw new IllegalArgumentException("Node does not belong to the same document which was rendered.");
        }

        if (node.getNodeType() == NodeType.DOCUMENT) {
            return new LayoutCollection<LayoutEntity>(mChildEntities);
        }

        java.util.ArrayList<LayoutEntity> entities = new java.util.ArrayList<LayoutEntity>();
        for (LayoutEntity entity : GetChildEntities(~LayoutEntityType.NONE, true)) {
            try {
                // Retrieve all entities from the layout document (inversion of LayoutEntityType.None).

                if (entity.getParentNode() == node) {
                    entities.add(entity);
                }

                if (entity.getType() == LayoutEntityType.ROW) {
                    RenderedRow row = (RenderedRow) entity;
                    if (row.getTable() == node) {
                        entities.add(entity);
                    }
                }

            } catch (RuntimeException ex) {

                throw ex;
            }
        }

        return new LayoutCollection<LayoutEntity>(entities);
    }

    private void ProcessLayoutElements(LayoutEntity current) throws Exception {
        do {
            LayoutEntity child = current.AddChildEntity(mEnumerator);

            if (mEnumerator.moveFirstChild()) {
                current = child;

                ProcessLayoutElements(current);
                mEnumerator.moveParent();

                current = current.getParent();
            }
        } while (mEnumerator.moveNext());
    }

    private void CollectLinesAndAddToMarkers() {
        CollectLinesOfMarkersCore((int) LayoutEntityType.COLUMN);
        CollectLinesOfMarkersCore((int) LayoutEntityType.COMMENT);
    }

    private void CollectLinesOfMarkersCore(int type) {
        java.util.ArrayList<RenderedLine> collectedLines = new java.util.ArrayList<RenderedLine>();

        for (RenderedPage page : getPages()) {
            //System.out.println(page.getText());
            for (LayoutEntity story : page.GetChildEntities(type, false)) {
                //RenderedLine
                for (LayoutEntity le : story.GetChildEntities((int) 32, true)) //LayoutEntityType.LINE
                {
                    RenderedLine line = (RenderedLine) le;
                    collectedLines.add(line);
                    for (RenderedSpan span : line.getSpans()) {
                        if (span.getKind().equals("PARAGRAPH") || span.getKind().equals("ROW") || span.getKind().equals("CELL") || span.getKind().equals("SECTION")) {
                            mLayoutToLinesLookup.put(span.getLayoutObject(), collectedLines);
                            collectedLines = new java.util.ArrayList<RenderedLine>();
                        } else {
                            mLayoutToSpanLookup.put(span.getLayoutObject(), span);
                        }
                    }
                }
            }
        }
    }

    @SuppressWarnings("unchecked")
    private void LinkLayoutMarkersToNodes(Document doc) throws Exception {
        for (Node node : (Iterable<Node>) doc.getChildNodes(NodeType.ANY, true)) {
            switch (node.getNodeType()) {
                case NodeType.PARAGRAPH:
                    for (RenderedLine line : GetLinesOfNode(node)) {
                        line.setParentNode(node);
                    }
                    break;

                case NodeType.ROW:
                    for (RenderedLine line : GetLinesOfNode(node)) {
                        line.setParentNode(((Row) node).getLastCell().getLastParagraph());
                    }
                    break;

                default:
                    if (mLayoutCollector.getEntity(node) != null) {
                        System.out.println(mLayoutCollector.getEntity(node));
                        //System.out.println(mLayoutToSpanLookup.get(mLayoutCollector.getEntity(node)));

                        mLayoutToSpanLookup.get(mLayoutCollector.getEntity(node)).setParentNode(node);
                    }
                    break;
            }
        }
    }

    private java.util.ArrayList<RenderedLine> GetLinesOfNode(Node node) throws Exception {
        java.util.ArrayList<RenderedLine> lines = new java.util.ArrayList<RenderedLine>();
        Object nodeEntity = mLayoutCollector.getEntity(node);

        if (nodeEntity != null && mLayoutToLinesLookup.containsKey(nodeEntity)) {
            lines = mLayoutToLinesLookup.get(nodeEntity);
        }

        return lines;
    }

    private LayoutCollector mLayoutCollector;
    private LayoutEnumerator mEnumerator;
    private static java.util.HashMap<Object, java.util.ArrayList<RenderedLine>> mLayoutToLinesLookup = new java.util.HashMap<Object, java.util.ArrayList<RenderedLine>>();
    private static java.util.HashMap<Object, RenderedSpan> mLayoutToSpanLookup = new java.util.HashMap<Object, RenderedSpan>();
}
