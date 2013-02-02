//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////


package Examples;

import com.aspose.words.*;
import org.testng.annotations.Test;

import java.util.Date;
import java.util.Locale;
import java.util.ArrayList;
import java.util.regex.Pattern;


public class ExField extends ExBase
{
    @Test
    public void updateTOC() throws Exception
    {
        Document doc = new Document();

        //ExStart
        //ExId:UpdateTOC
        //ExSummary:Shows how to completely rebuild TOC fields in the document by invoking field update.
        doc.updateFields();
        //ExEnd
    }

    @Test
    public void GetFieldType() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.TableOfContents.doc");

        //ExStart
        //ExFor:FieldType
        //ExFor:FieldChar
        //ExFor:FieldChar.FieldType
        //ExSummary:Shows how to find the type of field that is represented by a node which is derived from FieldChar.
        FieldChar fieldStart = (FieldChar)doc.getChild(NodeType.FIELD_START, 0, true);
        int type = fieldStart.getFieldType();
        //ExEnd
    }

    @Test
    public void insertTCField() throws Exception
    {
        //ExStart
        //ExId:InsertTCField
        //ExSummary:Shows how to insert a TC field into the document using DocumentBuilder.
        // Create a blank document.
        Document doc = new Document();

        // Create a document builder to insert content with.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TC field at the current document builder position.
        builder.insertField("TC \"Entry Text\" \\f t");
        //ExEnd
    }

    @Test
    public void changeLocale() throws Exception
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder b = new DocumentBuilder(doc);
        b.insertField("MERGEFIELD Date");

        //ExStart
        //ExId:ChangeCurrentCulture
        //ExSummary:Shows how to change the culture used in formatting fields during update.
        // Store the current culture so it can be set back once mail merge is complete.
        Locale currentCulture = Locale.getDefault();
        // Set to German language so dates and numbers are formatted using this culture during mail merge.
        Locale.setDefault(new Locale("de", "DE"));

        // Execute mail merge
        doc.getMailMerge().execute(new String[]{"Date"}, new Object[]{new Date()});

        // Restore the original culture.
        Locale.setDefault(currentCulture);
        //ExEnd

        doc.save(getMyDir() + "Field.ChangeLocale Out.doc");
    }

    @Test
    public void removeTOCFromDocumentCaller() throws Exception
    {
        removeTOCFromDocument();
    }

    //ExStart
    //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
    //ExId:RemoveTableOfContents
    //ExSummary:Demonstrates how to remove a specified TOC from a document.
    public void removeTOCFromDocument() throws Exception
    {
        // Open a document which contains a TOC.
        Document doc = new Document(getMyDir() + "Document.TableOfContents.doc");

        // Remove the first table of contents from the document.
        removeTableOfContents(doc, 0);

        // Save the output.
        doc.save(getMyDir() + "Document.TableOfContentsRemoveTOC Out.doc");
    }

    /**
     * Removes the specified table of contents field from the document.
     *
     * @param doc The document to remove the field from.
     * @param index The zero-based index of the TOC to remove.
     */
    static void removeTableOfContents(Document doc, int index) throws Exception
    {
        // Store the FieldStart nodes of TOC fields in the document for quick access.
        ArrayList fieldStarts = new ArrayList();
        // This is a list to store the nodes found inside the specified TOC. They will be removed
        // at thee end of this method.
        ArrayList nodeList = new ArrayList();

        for (FieldStart start : (Iterable<FieldStart>) doc.getChildNodes(NodeType.FIELD_START, true))
        {
            if (start.getFieldType() == FieldType.FIELD_TOC)
            {
                // Add all FieldStarts which are of type FieldTOC.
                fieldStarts.add(start);
            }
        }

        // Ensure the TOC specified by the passed index exists.
        if (index > fieldStarts.size() - 1)
            throw new ArrayIndexOutOfBoundsException("TOC index is out of range");

        boolean isRemoving = true;
        // Get the FieldStart of the specified TOC.
        Node currentNode = (Node)fieldStarts.get(index);

        while (isRemoving)
        {
            // It is safer to store these nodes and delete them all at once later.
            nodeList.add(currentNode);
            currentNode = currentNode.nextPreOrder(doc);

            // Once we encounter a FieldEnd node of type FieldTOC then we know we are at the end
            // of the current TOC and we can stop here.
            if (currentNode.getNodeType() == NodeType.FIELD_END)
            {
                FieldEnd fieldEnd = (FieldEnd)currentNode;
                if (fieldEnd.getFieldType() == FieldType.FIELD_TOC)
                    isRemoving = false;
            }
        }

        // Remove all nodes found in the specified TOC.
        for (Node node : (Iterable<Node>) nodeList)
        {
            node.remove();
        }
    }
    //ExEnd

    @Test
    //ExStart
    //ExId:TCFieldsRangeReplace
    //ExSummary:Shows how to find and insert a TC field at text in a document.
    public void insertTCFieldsAtText() throws Exception
    {
        Document doc = new Document();

        // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
        doc.getRange().replace(Pattern.compile("The Beginning"), new InsertTCFieldHandler("Chapter 1", "\\l 1"), false);
    }

    public class InsertTCFieldHandler implements IReplacingCallback
    {
        // Store the text and switches to be used for the TC fields.
        private String mFieldText;
        private String mFieldSwitches;

        /**
         * The switches to use for each TC field. Can be an empty string or null.
         */
        public InsertTCFieldHandler(String switches) throws Exception
        {
            this(null, switches);
        }

        /**
         * The display text and the switches to use for each TC field. Display text Can be an empty string or null.
         */
        public InsertTCFieldHandler(String text, String switches) throws Exception
        {
            mFieldText = text;
            mFieldSwitches = switches;
        }

        public int replacing(ReplacingArgs args) throws Exception
        {
            // Create a builder to insert the field.
            DocumentBuilder builder = new DocumentBuilder((Document)args.getMatchNode().getDocument());
            // Move to the first node of the match.
            builder.moveTo(args.getMatchNode());

            // If the user specified text to be used in the field as display text then use that, otherwise use the
            // match string as the display text.
            String insertText;

            if (!(mFieldText == null || "".equals(mFieldText)))
                insertText = mFieldText;
            else
                insertText = args.getMatch().group();

            // Insert the TC field before this node using the specified string as the display text and user defined switches.
            builder.insertField(java.text.MessageFormat.format("TC \"{0}\" {1}", insertText, mFieldSwitches));

            // We have done what we want so skip replacement.
            return ReplaceAction.SKIP;
        }
    }
    //ExEnd
}

