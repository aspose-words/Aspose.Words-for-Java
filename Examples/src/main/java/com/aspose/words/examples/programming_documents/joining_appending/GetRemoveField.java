package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.text.MessageFormat;


public class GetRemoveField {
    private static String gDataDir;

    public static void main(String[] args) throws Exception {

        //ExStart:GetRemoveField
        // The path to the documents directory.
        gDataDir = Utils.getDataDir(GetRemoveField.class);

        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(gDataDir + "TestFile.Source.doc");

        // Restart the page numbering on the start of the source document.
        srcDoc.getFirstSection().getPageSetup().setRestartPageNumbering(true);
        srcDoc.getFirstSection().getPageSetup().setPageStartingNumber(1);

        // Append the source document to the end of the destination document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // After joining the documents the NUMPAGE fields will now display the total number of pages which
        // is undesired behaviour. Call this method to fix them by replacing them with PAGEREF fields.
        convertNumPageFieldsToPageRef(dstDoc);

        // This needs to be called in order to update the new fields with page numbers.
        dstDoc.updatePageLayout();

        dstDoc.save(gDataDir + "TestFile.ConvertNumPageFields Out.doc");
        //ExEnd:GetRemoveField


        System.out.println("Documents appended successfully.");
    }

    //ExStart:convertNumPageFieldsToPageRef

    /**
     * Replaces all NUMPAGES fields in the document with PAGEREF fields. The replacement field displays the total number
     * of pages in the sub document instead of the total pages in the document.
     *
     * @param doc The combined document to process.
     */
    public static void convertNumPageFieldsToPageRef(Document doc) throws Exception {
        // This is the prefix for each bookmark which signals where page numbering restarts.
        // The underscore "_" at the start inserts this bookmark as hidden in MS Word.
        final String BOOKMARK_PREFIX = "_SubDocumentEnd";
        // Field name of the NUMPAGES field.
        final String NUM_PAGES_FIELD_NAME = "NUMPAGES";
        // Field name of the PAGEREF field.
        final String PAGE_REF_FIELD_NAME = "PAGEREF";

        // Create a new DocumentBuilder which is used to insert the bookmarks and replacement fields.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Defines the number of page restarts that have been encountered and therefore the number of "sub" documents
        // found within this document.
        int subDocumentCount = 0;

        // Iterate through all sections in the document.
        for (Section section : doc.getSections()) {
            // This section has it's page numbering restarted so we will treat this as the start of a sub document.
            // Any PAGENUM fields in this inner document must be converted to special PAGEREF fields to correct numbering.
            if (section.getPageSetup().getRestartPageNumbering()) {
                // Don't do anything if this is the first section in the document. This part of the code will insert the bookmark marking
                // the end of the previous sub document so therefore it is not applicable for first section in the document.
                if (!section.equals(doc.getFirstSection())) {
                    // Get the previous section and the last node within the body of that section.
                    Section prevSection = (Section) section.getPreviousSibling();
                    Node lastNode = prevSection.getBody().getLastChild();

                    // Use the DocumentBuilder to move to this node and insert the bookmark there.
                    // This bookmark represents the end of the sub document.
                    builder.moveTo(lastNode);
                    builder.startBookmark(BOOKMARK_PREFIX + subDocumentCount);
                    builder.endBookmark(BOOKMARK_PREFIX + subDocumentCount);

                    // Increase the subdocument count to insert the correct bookmarks.
                    subDocumentCount++;
                }
            }

            // The last section simply needs the ending bookmark to signal that it is the end of the current sub document.
            if (section.equals(doc.getLastSection())) {
                // Insert the bookmark at the end of the body of the last section.
                // Don't increase the count this time as we are just marking the end of the document.
                Node lastNode = doc.getLastSection().getBody().getLastChild();
                builder.moveTo(lastNode);
                builder.startBookmark(BOOKMARK_PREFIX + subDocumentCount);
                builder.endBookmark(BOOKMARK_PREFIX + subDocumentCount);
            }

            // Iterate through each NUMPAGES field in the section and replace the field with a PAGEREF field referring to the bookmark of the current subdocument
            // This bookmark is positioned at the end of the sub document but does not exist yet. It is inserted when a section with restart page numbering or the last
            // section is encountered.
            for (Node node : section.getChildNodes(NodeType.FIELD_START, true).toArray()) {
                FieldStart fieldStart = (FieldStart) node;

                if (fieldStart.getFieldType() == FieldType.FIELD_NUM_PAGES) {
                    // Get the field code.
                    String fieldCode = getFieldCode(fieldStart);
                    // Since the NUMPAGES field does not take any additional parameters we can assume the remaining part of the field
                    // code after the fieldname are the switches. We will use these to help recreate the NUMPAGES field as a PAGEREF field.
                    String fieldSwitches = fieldCode.replace(NUM_PAGES_FIELD_NAME, "").trim();

                    // Inserting the new field directly at the FieldStart node of the original field will cause the new field to
                    // not pick up the formatting of the original field. To counter this insert the field just before the original field
                    Node previousNode = fieldStart.getPreviousSibling();

                    // If a previous run cannot be found then we are forced to use the FieldStart node.
                    if (previousNode == null)
                        previousNode = fieldStart;

                    // Insert a PAGEREF field at the same position as the field.
                    builder.moveTo(previousNode);
                    // This will insert a new field with a code like " PAGEREF _SubDocumentEnd0 *\MERGEFORMAT ".
                    Field newField = builder.insertField(MessageFormat.format(" {0} {1}{2} {3} ", PAGE_REF_FIELD_NAME, BOOKMARK_PREFIX, subDocumentCount, fieldSwitches));

                    // The field will be inserted before the referenced node. Move the node before the field instead.
                    previousNode.getParentNode().insertBefore(previousNode, newField.getStart());

                    // Remove the original NUMPAGES field from the document.
                    removeField(fieldStart);
                }
            }
        }
    }
    //ExEnd:convertNumPageFieldsToPageRef

    /**
     * Retrieves the field code from a field.
     *
     * @param fieldStart The field start of the field which to gather the field code from.
     */


    //ExStart:removeField

    /**
     * Removes the Field from the document.
     *
     * @param fieldStart The field start node of the field to remove.
     */


    private static void removeField(FieldStart fieldStart) throws Exception {
        Node currentNode = fieldStart;
        boolean isRemoving = true;
        while (currentNode != null && isRemoving) {
            if (currentNode.getNodeType() == NodeType.FIELD_END)
                isRemoving = false;

            Node nextNode = currentNode.nextPreOrder(currentNode.getDocument());
            currentNode.remove();
            currentNode = nextNode;
        }
    }
    //ExEnd:removeField

    //ExStart:getFieldCode
    private static String getFieldCode(FieldStart fieldStart) throws Exception {
        StringBuilder builder = new StringBuilder();

        for (Node node = fieldStart; node != null && node.getNodeType() != NodeType.FIELD_SEPARATOR &&
                node.getNodeType() != NodeType.FIELD_END; node = node.nextPreOrder(node.getDocument())) {
            // Use text only of Run nodes to avoid duplication.
            if (node.getNodeType() == NodeType.RUN)
                builder.append(node.getText());
        }
        return builder.toString();
    }
    //ExEnd:getFieldCode


}