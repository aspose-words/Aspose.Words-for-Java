/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package programmingwithdocuments.joiningandappending.appenddocument.java;

import java.text.MessageFormat;
import java.util.Arrays;
import java.util.HashMap;
import java.util.ArrayList;
import java.util.Collections;
import java.io.File;
import java.net.URI;

import com.aspose.words.*;


public class AppendDocument
{
    private static String gDataDir;

    public static void main(String[] args) throws Exception
    {
            // The path to the documents directory.
        gDataDir = "src/programmingwithdocuments/joiningandappending/appenddocument/data/";

        // Run each of the sample code snippets.
        appendDocument_SimpleAppendDocument();
        appendDocument_KeepSourceFormatting();
        appendDocument_UseDestinationStyles();
        appendDocument_JoinContinuous();
        appendDocument_JoinNewPage();
        appendDocument_RestartPageNumbering();
        appendDocument_LinkHeadersFooters();
        appendDocument_UnlinkHeadersFooters();
        appendDocument_RemoveSourceHeadersFooters();
        appendDocument_DifferentPageSetup();
        appendDocument_ConvertNumPageFields();
        appendDocument_ListUseDestinationStyles();
        appendDocument_ListKeepSourceFormatting();
        appendDocument_KeepSourceTogether();
        appendDocument_BaseDocument();
        appendDocument_UpdatePageLayout();
    }

    public static void appendDocument_SimpleAppendDocument() throws Exception
    {
        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

        //ExStart
        //ExId:AppendDocument_SimpleAppend
        //ExSummary:Shows how to append a document to the end of another document using no additional options.
        // Append the source document to the destination document using no extra options.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        //ExEnd

        dstDoc.save(gDataDir + "TestFile.SimpleAppendDocument Out.docx");
    }

    public static void appendDocument_KeepSourceFormatting() throws Exception
    {
        //ExStart
        //ExId:AppendDocument_KeepSourceFormatting
        //ExSummary:Shows how to append a document to another document while keeping the original formatting.
        // Load the documents to join.
        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

        // Keep the formatting from the source document when appending it to the destination document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // Save the joined document to disk.
        dstDoc.save(gDataDir + "TestFile.KeepSourceFormatting Out.docx");
        //ExEnd
    }

    public static void appendDocument_UseDestinationStyles() throws Exception
    {
        //ExStart
        //ExId:AppendDocument_UseDestinationStyles
        //ExSummary:Shows how to append a document to another document using the formatting of the destination document.
        // Load the documents to join.
        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

        // Append the source document using the styles of the destination document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES);

        // Save the joined document to disk.
        dstDoc.save(gDataDir + "TestFile.UseDestinationStyles Out.doc");
        //ExEnd
    }

    public static void appendDocument_JoinContinuous() throws Exception
    {
        //ExStart
        //ExId:AppendDocument_JoinContinuous
        //ExSummary:Shows how to append a document to another document so the content flows continuously.
        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

        // Make the document appear straight after the destination documents content.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        // Append the source document using the original styles found in the source document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(gDataDir + "TestFile.JoinContinuous Out.doc");
        //ExEnd
    }

    public static void appendDocument_JoinNewPage() throws Exception
    {
        //ExStart
        //ExId:AppendDocument_JoinNewPage
        //ExSummary:Shows how to append a document to another document so it starts on a new page.
        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

        // Set the appended document to start on a new page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.NEW_PAGE);

        // Append the source document using the original styles found in the source document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(gDataDir + "TestFile.JoinNewPage Out.doc");
        //ExEnd
    }

    public static void appendDocument_RestartPageNumbering() throws Exception
    {
        //ExStart
        //ExId:AppendDocument_RestartPageNumbering
        //ExSummary:Shows how to append a document to another document with page numbering restarted.
        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

        // Set the appended document to appear on the next page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.NEW_PAGE);
        // Restart the page numbering for the document to be appended.
        srcDoc.getFirstSection().getPageSetup().setRestartPageNumbering(true);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(gDataDir + "TestFile.RestartPageNumbering Out.doc");
        //ExEnd
    }

    public static void appendDocument_LinkHeadersFooters() throws Exception
    {
        //ExStart
        //ExFor:HeaderFooterCollection.LinkToPrevious(Boolean)
        //ExId:AppendDocument_LinkHeadersFooters
        //ExSummary:Shows how to append a document to another document and continue headers and footers from the destination document.
        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

        // Set the appended document to appear on a new page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.NEW_PAGE);

        // Link the headers and footers in the source document to the previous section.
        // This will override any headers or footers already found in the source document.
        srcDoc.getFirstSection().getHeadersFooters().linkToPrevious(true);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(gDataDir + "TestFile.LinkHeadersFooters Out.doc");
        //ExEnd
    }

    public static void appendDocument_UnlinkHeadersFooters() throws Exception
    {
        //ExStart
        //ExId:AppendDocument_UnlinkHeadersFooters
        //ExSummary:Shows how to append a document to another document so headers and footers do not continue from the destination document.
        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

        // Even a document with no headers or footers can still have the LinkToPrevious setting set to true.
        // Unlink the headers and footers in the source document to stop this from continuing the headers and footers
        // from the destination document.
        srcDoc.getFirstSection().getHeadersFooters().linkToPrevious(false);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(gDataDir + "TestFile.UnlinkHeadersFooters Out.doc");
        //ExEnd
    }

    public static void appendDocument_RemoveSourceHeadersFooters() throws Exception
    {
        //ExStart
        //ExId:AppendDocument_RemoveSourceHeadersFooters
        //ExSummary:Shows how to remove headers and footers from a document before appending it to another document.
        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

        // Remove the headers and footers from each of the sections in the source document.
        for (Section section : srcDoc.getSections())
        {
            section.clearHeadersFooters();
        }

        // Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting
        // for HeadersFooters can still be set. This will cause the headers and footers to continue from the destination
        // document. This should set to false to avoid this behaviour.
        srcDoc.getFirstSection().getHeadersFooters().linkToPrevious(false);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(gDataDir + "TestFile.RemoveSourceHeadersFooters Out.doc");
        //ExEnd
    }

    public static void appendDocument_DifferentPageSetup() throws Exception
    {
        //ExStart
        //ExId:AppendDocument_DifferentPageSetup
        //ExSummary:Shows how to append a document to another document continuously which has different page settings.
        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.SourcePageSetup.doc");

        // Set the source document to continue straight after the end of the destination document.
        // If some page setup settings are different then this may not work and the source document will appear
        // on a new page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        // To ensure this does not happen when the source document has different page setup settings make sure the
        // settings are identical between the last section of the destination document.
        // If there are further continuous sections that follow on in the source document then this will need to be
        // repeated for those sections as well.
        srcDoc.getFirstSection().getPageSetup().setPageWidth(dstDoc.getLastSection().getPageSetup().getPageWidth());
        srcDoc.getFirstSection().getPageSetup().setPageHeight(dstDoc.getLastSection().getPageSetup().getPageHeight());
        srcDoc.getFirstSection().getPageSetup().setOrientation(dstDoc.getLastSection().getPageSetup().getOrientation());

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(gDataDir + "TestFile.DifferentPageSetup Out.doc");
        //ExEnd
    }

    //ExStart
    //ExId:AppendDocument_ConvertNumPageFields
    //ExSummary:Shows how to change the NUMPAGE fields in a document to display the number of pages only within a sub document.
    public static void appendDocument_ConvertNumPageFields() throws Exception
    {
        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

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
    }

    /**
     * Replaces all NUMPAGES fields in the document with PAGEREF fields. The replacement field displays the total number
     * of pages in the sub document instead of the total pages in the document.
     *
     * @param doc The combined document to process.
     */
    public static void convertNumPageFieldsToPageRef(Document doc) throws Exception
    {
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
        for (Section section : doc.getSections())
        {
            // This section has it's page numbering restarted so we will treat this as the start of a sub document.
            // Any PAGENUM fields in this inner document must be converted to special PAGEREF fields to correct numbering.
            if (section.getPageSetup().getRestartPageNumbering())
            {
                // Don't do anything if this is the first section in the document. This part of the code will insert the bookmark marking
                // the end of the previous sub document so therefore it is not applicable for first section in the document.
                if (!section.equals(doc.getFirstSection()))
                {
                    // Get the previous section and the last node within the body of that section.
                    Section prevSection = (Section)section.getPreviousSibling();
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
            if (section.equals(doc.getLastSection()))
            {
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
            for (Node node : section.getChildNodes(NodeType.FIELD_START, true).toArray())
            {
                FieldStart fieldStart = (FieldStart)node;

                if (fieldStart.getFieldType() == FieldType.FIELD_NUM_PAGES)
                {
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
    //ExEnd

    public static void appendDocument_ListUseDestinationStyles() throws Exception
    {
        //ExStart
        //ExId:AppendDocument_ListUseDestinationStyles
        //ExSummary:Shows how to append a document using destination styles and preventing any list numberings from continuing on.
        Document dstDoc = new Document(gDataDir + "TestFile.DestinationList.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.SourceList.doc");

        // Set the source document to continue straight after the end of the destination document.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        // Keep track of the lists that are created.
        HashMap newLists = new HashMap();

        // Iterate through all paragraphs in the document.
        for (Paragraph para : (Iterable<Paragraph>) srcDoc.getChildNodes(NodeType.PARAGRAPH, true))
        {
            if (para.isListItem())
            {
                int listId = para.getListFormat().getList().getListId();

                // Check if the destination document contains a list with this ID already. If it does then this may
                // cause the two lists to run together. Create a copy of the list in the source document instead.
                if (dstDoc.getLists().getListByListId(listId) != null)
                {
                    List currentList;
                    // A newly copied list already exists for this ID, retrieve the stored list and use it on
                    // the current paragraph.
                    if (newLists.containsKey(listId))
                    {
                        currentList = (List)newLists.get(listId);
                    }
                    else
                    {
                        // Add a copy of this list to the document and store it for later reference.
                        currentList = srcDoc.getLists().addCopy(para.getListFormat().getList());
                        newLists.put(listId, currentList);
                    }

                    // Set the list of this paragraph  to the copied list.
                    para.getListFormat().setList(currentList);
                }
            }
        }

        // Append the source document to end of the destination document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES);

        // Save the combined document to disk.
        dstDoc.save(gDataDir + "TestFile.ListUseDestinationStyles Out.docx");
        //ExEnd
    }

    public static void appendDocument_ListKeepSourceFormatting() throws Exception
    {
        //ExStart
        //ExId:AppendDocument_ListKeepSourceFormatting
        //ExSummary:Shows how to append a document to another document containing lists retaining source formatting.
        Document dstDoc = new Document(gDataDir + "TestFile.DestinationList.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.SourceList.doc");

        // Append the content of the document so it flows continuously.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(gDataDir + "TestFile.ListKeepSourceFormatting Out.doc");
        //ExEnd
    }

    public static void appendDocument_KeepSourceTogether() throws Exception
    {
        //ExStart
        //ExFor:ParagraphFormat.KeepWithNext
        //ExId:AppendDocument_KeepSourceTogether
        //ExSummary:Shows how to append a document to another document while keeping the content from splitting across two pages.
        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

        // Set the source document to appear straight after the destination document's content.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        // Iterate through all sections in the source document.
        for(Paragraph para : (Iterable<Paragraph>) srcDoc.getChildNodes(NodeType.PARAGRAPH, true))
        {
            para.getParagraphFormat().setKeepWithNext(true);
        }

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(gDataDir + "TestDcc.KeepSourceTogether Out.doc");
        //ExEnd
    }

    public static void appendDocument_BaseDocument() throws Exception
    {
        //ExStart
        //ExId:AppendDocument_BaseDocument
        //ExSummary:Shows how to remove all content from a document before using it as a base to append documents to.
        // Use a blank document as the destination document.
        Document dstDoc = new Document();
        Document srcDoc = new Document(gDataDir + "TestFile.Source.doc");

        // The destination document is not actually empty which often causes a blank page to appear before the appended document
        // This is due to the base document having an empty section and the new document being started on the next page.
        // Remove all content from the destination document before appending.
        dstDoc.removeAllChildren();

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(gDataDir + "TestFile.BaseDocument Out.doc");
        //ExEnd
    }

    public static void appendDocument_UpdatePageLayout() throws Exception
    {
        //ExStart
        //ExId:AppendDocument_UpdatePageLayout
        //ExSummary:Shows how to rebuild the document layout after appending further content.
        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(gDataDir + "TestFile.Source.doc");

        // If the destination document is rendered to PDF, image etc or UpdatePageLayout is called before the source document
        // is appended then any changes made after will not be reflected in the rendered output.
        dstDoc.updatePageLayout();

        // Join the documents.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // For the changes to be updated to rendered output, UpdatePageLayout must be called again.
        // If not called again the appended document will not appear in the output of the next rendering.
        dstDoc.updatePageLayout();

        // Save the joined document to PDF.
        dstDoc.save(gDataDir + "TestFile.UpdatePageLayout Out.pdf");
        //ExEnd
    }

    //ExStart
    //ExFor:FieldStart
	//ExFor:FieldSeparator
    //ExFor:FieldEnd
    //ExId:AppendDocument_HelperFunctions
    //ExSummary:Provides some helper functions by the methods above
    /**
     * Retrieves the field code from a field.
     *
     * @param fieldStart The field start of the field which to gather the field code from.
     */
    private static String getFieldCode(FieldStart fieldStart) throws Exception
    {
        StringBuilder builder = new StringBuilder();

        for (Node node = fieldStart; node != null && node.getNodeType() != NodeType.FIELD_SEPARATOR &&
            node.getNodeType() != NodeType.FIELD_END; node = node.nextPreOrder(node.getDocument()))
        {
            // Use text only of Run nodes to avoid duplication.
            if (node.getNodeType() == NodeType.RUN)
                builder.append(node.getText());
        }
        return builder.toString();
    }

    /**
     * Removes the Field from the document.
     *
     * @param fieldStart The field start node of the field to remove.
     */
    private static void removeField(FieldStart fieldStart) throws Exception
    {
        Node currentNode = fieldStart;
        boolean isRemoving = true;
        while (currentNode != null && isRemoving)
        {
            if (currentNode.getNodeType() == NodeType.FIELD_END)
                isRemoving = false;

            Node nextNode = currentNode.nextPreOrder(currentNode.getDocument());
            currentNode.remove();
            currentNode = nextNode;
        }
    }
    //ExEnd

    //ExStart
    //ExFor:DocumentBase.ImportNode(Node,bool,ImportFormatMode)
    //ExFor:ImportFormatMode
    //ExId:CombineDocuments
    //ExSummary:Shows how to manually append the content from one document to the end of another document.
    /**
     * A manual implementation of the Document.AppendDocument function which shows the general
     * steps of how a document is appended to another.
     *
     * @param dstDoc The destination document where to append to.
     * @param srcDoc The source document.
     * @param mode The import mode to use when importing content from another document.
     */
    public void appendDocument(Document dstDoc, Document srcDoc, int mode) throws Exception
    {
        // Loop through all sections in the source document.
        // Section nodes are immediate children of the Document node so we can just enumerate the Document.
        for (Node srcNode : srcDoc)
        {
            Section srcSection = (Section)srcNode;

            // Because we are copying a section from one document to another,
            // it is required to import the Section node into the destination document.
            // This adjusts any document-specific references to styles, lists, etc.
            //
            // Importing a node creates a copy of the original node, but the copy
            // is ready to be inserted into the destination document.
            Node dstSection = dstDoc.importNode(srcSection, true, mode);

            // Now the new section node can be appended to the destination document.
            dstDoc.appendChild(dstSection);
        }
    }
    //ExEnd

    //ExStart
    //ExFor:DocumentBase.ImportNode(Node,bool,ImportFormatMode)
    //ExFor:CompositeNode.PrependChild(Node)
    //ExFor:ImportFormatMode
    //ExId:PrependDocument
    //ExSummary:Shows how to manually prepend the content from one document to the beginning of another document.
    public static void prependDocumentMain() throws Exception
    {
        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(gDataDir + "TestFile.Source.doc");

        // Append the source document to the destination document. This causes the result to have line spacing problems.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // Instead prepend the content of the destination document to the start of the source document.
        // This results in the same joined document but with no line spacing issues.
        prependDocument(srcDoc, dstDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
    }


    /**
     * A modified version of the AppendDocument method which prepends the content of one document to the start
     * of another.
     *
     * @param dstDoc The destination document where to prepend the source document to.
     * @param srcDoc The source document.
     */
    public static void prependDocument(Document dstDoc, Document srcDoc, int mode) throws Exception
    {
        // Loop through all sections in the source document.
        // Section nodes are immediate children of the Document node so we can just enumerate the Document.
        ArrayList sections = (ArrayList)Arrays.asList(srcDoc.getSections().toArray());

        // Reverse the order of the sections so they are prepended to start of the destination document in the correct order.
        Collections.reverse(sections);

        for (Section srcSection : (Iterable<Section>) sections)
        {
            // Import the nodes from the source document.
            Node dstSection = dstDoc.importNode(srcSection, true, mode);

            // Now the new section node can be prepended to the destination document.
            // Note how PrependChild is used instead of AppendChild. This is the only line changed compared
            // to the original method.
            dstDoc.prependChild(dstSection);
        }
    }
    //ExEnd

}