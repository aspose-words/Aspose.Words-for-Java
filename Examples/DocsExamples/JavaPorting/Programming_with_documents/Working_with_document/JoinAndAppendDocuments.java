package DocsExamples.Programming_with_Documents.Working_with_Document;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.Section;
import com.aspose.words.Node;
import com.aspose.words.ImportFormatOptions;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.NodeType;
import com.aspose.words.FieldStart;
import com.aspose.words.FieldType;
import com.aspose.words.Field;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.words.SectionStart;
import com.aspose.words.Paragraph;
import java.util.HashMap;
import com.aspose.words.List;
import com.aspose.ms.System.Collections.msDictionary;
import com.aspose.words.BreakType;
import com.aspose.words.NodeImporter;
import com.aspose.words.ParagraphCollection;


class JoinAndAppendDocuments extends DocsExamplesBase
{
    @Test
    public void simpleAppendDocument() throws Exception
    {
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Northwind traders.docx");

        // Append the source document to the destination document using no extra options.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.SimpleAppendDocument.docx");
    }

    @Test
    public void appendDocument() throws Exception
    {
        //ExStart:AppendDocumentManually
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Northwind traders.docx");
        
        // Loop through all sections in the source document.
        // Section nodes are immediate children of the Document node so we can just enumerate the Document.
        for (Section srcSection : (Iterable<Section>) srcDoc)
        {
            // Because we are copying a section from one document to another, 
            // it is required to import the Section node into the destination document.
            // This adjusts any document-specific references to styles, lists, etc.
            //
            // Importing a node creates a copy of the original node, but the copy
            // ss ready to be inserted into the destination document.
            Node dstSection = dstDoc.importNode(srcSection, true, ImportFormatMode.KEEP_SOURCE_FORMATTING);

            // Now the new section node can be appended to the destination document.
            dstDoc.appendChild(dstSection);
        }

        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.AppendDocument.docx");
        //ExEnd:AppendDocumentManually
    }

    @Test
    public void appendDocumentToBlank() throws Exception
    {
        //ExStart:AppendDocumentToBlank
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document();
        
        // The destination document is not empty, often causing a blank page to appear before the appended document.
        // This is due to the base document having an empty section and the new document being started on the next page.
        // Remove all content from the destination document before appending.
        dstDoc.removeAllChildren();
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        
        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.AppendDocumentToBlank.docx");
        //ExEnd:AppendDocumentToBlank
    }

    @Test
    public void appendWithImportFormatOptions() throws Exception
    {
        //ExStart:AppendWithImportFormatOptions
        Document srcDoc = new Document(getMyDir() + "Document source with list.docx");
        Document dstDoc = new Document(getMyDir() + "Document destination with list.docx");

        // Specify that if numbering clashes in source and destination documents,
        // then numbering from the source document will be used.
        ImportFormatOptions options = new ImportFormatOptions(); { options.setKeepSourceNumbering(true); }
        
        dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES, options);
        //ExEnd:AppendWithImportFormatOptions
    }

    @Test
    public void convertNumPageFields() throws Exception
    {
        //ExStart:ConvertNumPageFields
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Northwind traders.docx");

        // Restart the page numbering on the start of the source document.
        srcDoc.getFirstSection().getPageSetup().setRestartPageNumbering(true);
        srcDoc.getFirstSection().getPageSetup().setPageStartingNumber(1);

        // Append the source document to the end of the destination document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // After joining the documents the NUMPAGE fields will now display the total number of pages which
        // is undesired behavior. Call this method to fix them by replacing them with PAGEREF fields.
        convertNumPageFieldsToPageRef(dstDoc);

        // This needs to be called in order to update the new fields with page numbers.
        dstDoc.updatePageLayout();

        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.ConvertNumPageFields.docx");
        //ExEnd:ConvertNumPageFields
    }

    //ExStart:ConvertNumPageFieldsToPageRef
    public void convertNumPageFieldsToPageRef(Document doc) throws Exception
    {
        // This is the prefix for each bookmark, which signals where page numbering restarts.
        // The underscore "_" at the start inserts this bookmark as hidden in MS Word.
        final String BOOKMARK_PREFIX = "_SubDocumentEnd";
        final String NUM_PAGES_FIELD_NAME = "NUMPAGES";
        final String PAGE_REF_FIELD_NAME = "PAGEREF";

        // Defines the number of page restarts encountered and, therefore,
        // the number of "sub" documents found within this document.
        int subDocumentCount = 0;

        DocumentBuilder builder = new DocumentBuilder(doc);
        
        for (Section section : (Iterable<Section>) doc.getSections())
        {
            // This section has its page numbering restarted to treat this as the start of a sub-document.
            // Any PAGENUM fields in this inner document must be converted to special PAGEREF fields to correct numbering.
            if (section.getPageSetup().getRestartPageNumbering())
            {
                // Don't do anything if this is the first section of the document.
                // This part of the code will insert the bookmark marking the end of the previous sub-document so,
                // therefore, it does not apply to the first section in the document.
                if (!section.equals(doc.getFirstSection()))
                {
                    // Get the previous section and the last node within the body of that section.
                    Section prevSection = (Section) section.getPreviousSibling();
                    Node lastNode = prevSection.getBody().getLastChild();

                    builder.moveTo(lastNode);
                    
                    // This bookmark represents the end of the sub-document.
                    builder.startBookmark(BOOKMARK_PREFIX + subDocumentCount);
                    builder.endBookmark(BOOKMARK_PREFIX + subDocumentCount);

                    // Increase the sub-document count to insert the correct bookmarks.
                    subDocumentCount++;
                }
            }

            // The last section needs the ending bookmark to signal that it is the end of the current sub-document.
            if (section.equals(doc.getLastSection()))
            {
                // Insert the bookmark at the end of the body of the last section.
                // Don't increase the count this time as we are just marking the end of the document.
                Node lastNode = doc.getLastSection().getBody().getLastChild();
                
                builder.moveTo(lastNode);
                builder.startBookmark(BOOKMARK_PREFIX + subDocumentCount);
                builder.endBookmark(BOOKMARK_PREFIX + subDocumentCount);
            }

            // Iterate through each NUMPAGES field in the section and replace it with a PAGEREF field
            // referring to the bookmark of the current sub-document. This bookmark is positioned at the end
            // of the sub-document but does not exist yet. It is inserted when a section with restart page numbering
            // or the last section is encountered.
            Node[] nodes = section.getChildNodes(NodeType.FIELD_START, true).toArray();
            
            for (FieldStart fieldStart : nodes)
            {
                if (fieldStart.getFieldType() == FieldType.FIELD_NUM_PAGES)
                {
                    String fieldCode = getFieldCode(fieldStart);
                    // Since the NUMPAGES field does not take any additional parameters,
                    // we can assume the field's remaining part. Code after the field name is the switches.
                    // We will use these to help recreate the NUMPAGES field as a PAGEREF field.
                    String fieldSwitches = fieldCode.replace(NUM_PAGES_FIELD_NAME, "").trim();

                    // Inserting the new field directly at the FieldStart node of the original field will cause
                    // the new field not to pick up the original field's formatting. To counter this,
                    // insert the field just before the original field if a previous run cannot be found,
                    // we are forced to use the FieldStart node.
                    Node previousNode = (fieldStart.getPreviousSibling() != null ? fieldStart.getPreviousSibling() : fieldStart);
                    
                    // Insert a PAGEREF field at the same position as the field.
                    builder.moveTo(previousNode);
                    
                    Field newField = builder.insertField(
                        $" {pageRefFieldName} {bookmarkPrefix}{subDocumentCount} {fieldSwitches} ");

                    // The field will be inserted before the referenced node. Move the node before the field instead.
                    previousNode.getParentNode().insertBefore(previousNode, newField.getStart());

                    // Remove the original NUMPAGES field from the document.
                    removeField(fieldStart);
                }
            }
        }
    }
    //ExEnd:ConvertNumPageFieldsToPageRef
    
    //ExStart:GetRemoveField
    private void removeField(FieldStart fieldStart)
    {
        boolean isRemoving = true;
        
        Node currentNode = fieldStart;
        while (currentNode != null && isRemoving)
        {
            if (currentNode.getNodeType() == NodeType.FIELD_END)
                isRemoving = false;

            Node nextNode = currentNode.nextPreOrder(currentNode.getDocument());
            currentNode.remove();
            currentNode = nextNode;
        }
    }

    private String getFieldCode(FieldStart fieldStart)
    {
        StringBuilder builder = new StringBuilder();

        for (Node node = fieldStart;
            node != null && node.getNodeType() != NodeType.FIELD_SEPARATOR &&
            node.getNodeType() != NodeType.FIELD_END;
            node = node.nextPreOrder(node.getDocument()))
        {
            // Use text only of Run nodes to avoid duplication.
            if (node.getNodeType() == NodeType.RUN)
                msStringBuilder.append(builder, node.getText());
        }

        return builder.toString();
    }
    //ExEnd:GetRemoveField

    @Test
    public void differentPageSetup() throws Exception
    {
        //ExStart:DifferentPageSetup
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Northwind traders.docx");

        // Set the source document to continue straight after the end of the destination document.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        // Restart the page numbering on the start of the source document.
        srcDoc.getFirstSection().getPageSetup().setRestartPageNumbering(true);
        srcDoc.getFirstSection().getPageSetup().setPageStartingNumber(1);

        // To ensure this does not happen when the source document has different page setup settings, make sure the
        // settings are identical between the last section of the destination document.
        // If there are further continuous sections that follow on in the source document,
        // this will need to be repeated for those sections.
        srcDoc.getFirstSection().getPageSetup().setPageWidth(dstDoc.getLastSection().getPageSetup().getPageWidth());
        srcDoc.getFirstSection().getPageSetup().setPageHeight(dstDoc.getLastSection().getPageSetup().getPageHeight());
        srcDoc.getFirstSection().getPageSetup().setOrientation(dstDoc.getLastSection().getPageSetup().getOrientation());

        // Iterate through all sections in the source document.
        for (Paragraph para : (Iterable<Paragraph>) srcDoc.getChildNodes(NodeType.PARAGRAPH, true))
        {
            para.getParagraphFormat().setKeepWithNext(true);
        }

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        
        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.DifferentPageSetup.docx");
        //ExEnd:DifferentPageSetup
    }

    @Test
    public void joinContinuous() throws Exception
    {
        //ExStart:JoinContinuous
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Northwind traders.docx");

        // Make the document appear straight after the destination documents content.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);
        // Append the source document using the original styles found in the source document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        
        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.JoinContinuous.docx");
        //ExEnd:JoinContinuous
    }

    @Test
    public void joinNewPage() throws Exception
    {
        //ExStart:JoinNewPage
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Northwind traders.docx");

        // Set the appended document to start on a new page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.NEW_PAGE);
        // Append the source document using the original styles found in the source document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        
        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.JoinNewPage.docx");
        //ExEnd:JoinNewPage
    }

    @Test
    public void keepSourceFormatting() throws Exception
    {
        //ExStart:KeepSourceFormatting
        Document dstDoc = new Document();
        dstDoc.getFirstSection().getBody().appendParagraph("Destination document text. ");

        Document srcDoc = new Document();
        srcDoc.getFirstSection().getBody().appendParagraph("Source document text. ");

        // Append the source document to the destination document.
        // Pass format mode to retain the original formatting of the source document when importing it.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.KeepSourceFormatting.docx");
        //ExEnd:KeepSourceFormatting
    }

    @Test
    public void keepSourceTogether() throws Exception
    {
        //ExStart:KeepSourceTogether
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Document destination with list.docx");
        
        // Set the source document to appear straight after the destination document's content.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        for (Paragraph para : (Iterable<Paragraph>) srcDoc.getChildNodes(NodeType.PARAGRAPH, true))
        {
            para.getParagraphFormat().setKeepWithNext(true);
        }

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        
        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.KeepSourceTogether.docx");
        //ExEnd:KeepSourceTogether
    }        

    @Test
    public void listKeepSourceFormatting() throws Exception
    {
        //ExStart:ListKeepSourceFormatting
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Document destination with list.docx");

        // Append the content of the document so it flows continuously.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        
        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.ListKeepSourceFormatting.docx");
        //ExEnd:ListKeepSourceFormatting
    }

    @Test
    public void listUseDestinationStyles() throws Exception
    {
        //ExStart:ListUseDestinationStyles
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Document destination with list.docx");

        // Set the source document to continue straight after the end of the destination document.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        // Keep track of the lists that are created.
        HashMap<Integer, List> newLists = new HashMap<Integer, List>();

        for (Paragraph para : (Iterable<Paragraph>) srcDoc.getChildNodes(NodeType.PARAGRAPH, true))
        {
            if (para.isListItem())
            {
                int listId = para.getListFormat().getList().getListId();

                // Check if the destination document contains a list with this ID already. If it does, then this may
                // cause the two lists to run together. Create a copy of the list in the source document instead.
                if (dstDoc.getLists().getListByListId(listId) != null)
                {
                    List currentList;
                    // A newly copied list already exists for this ID, retrieve the stored list,
                    // and use it on the current paragraph.
                    if (newLists.containsKey(listId))
                    {
                        currentList = newLists.get(listId);
                    }
                    else
                    {
                        // Add a copy of this list to the document and store it for later reference.
                        currentList = srcDoc.getLists().addCopy(para.getListFormat().getList());
                        msDictionary.add(newLists, listId, currentList);
                    }

                    // Set the list of this paragraph to the copied list.
                    para.getListFormat().setList(currentList);
                }
            }
        }

        // Append the source document to end of the destination document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES);

        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.ListUseDestinationStyles.docx");
        //ExEnd:ListUseDestinationStyles
    }

    @Test
    public void restartPageNumbering() throws Exception
    {
        //ExStart:RestartPageNumbering
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Northwind traders.docx");

        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.NEW_PAGE);
        srcDoc.getFirstSection().getPageSetup().setRestartPageNumbering(true);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        
        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.RestartPageNumbering.docx");
        //ExEnd:RestartPageNumbering
    }

    @Test
    public void updatePageLayout() throws Exception
    {
        //ExStart:UpdatePageLayout
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Northwind traders.docx");

        // If the destination document is rendered to PDF, image etc.
        // or UpdatePageLayout is called before the source document. Is appended,
        // then any changes made after will not be reflected in the rendered output
        dstDoc.updatePageLayout();

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // For the changes to be updated to rendered output, UpdatePageLayout must be called again.
        // If not called again, the appended document will not appear in the output of the next rendering.
        dstDoc.updatePageLayout();

        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.UpdatePageLayout.docx");
        //ExEnd:UpdatePageLayout
    }

    @Test
    public void useDestinationStyles() throws Exception
    {
        //ExStart:UseDestinationStyles
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Northwind traders.docx");

        // Append the source document using the styles of the destination document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES);

        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.UseDestinationStyles.docx");
        //ExEnd:UseDestinationStyles
    }

    @Test
    public void smartStyleBehavior() throws Exception
    {
        //ExStart:SmartStyleBehavior
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Northwind traders.docx");
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        
        builder.moveToDocumentEnd();
        builder.insertBreak(BreakType.PAGE_BREAK);

        ImportFormatOptions options = new ImportFormatOptions(); { options.setSmartStyleBehavior(true); }

        builder.insertDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES, options);
        builder.getDocument().save(getArtifactsDir() + "JoinAndAppendDocuments.SmartStyleBehavior.docx");
        //ExEnd:SmartStyleBehavior
    }

    @Test
    public void insertDocumentWithBuilder() throws Exception
    {
        //ExStart:InsertDocumentWithBuilder
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Northwind traders.docx");
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        builder.moveToDocumentEnd();
        builder.insertBreak(BreakType.PAGE_BREAK);

        builder.insertDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        builder.getDocument().save(getArtifactsDir() + "JoinAndAppendDocuments.InsertDocumentWithBuilder.docx");
        //ExEnd:InsertDocumentWithBuilder
    }

    @Test
    public void keepSourceNumbering() throws Exception
    {
        //ExStart:KeepSourceNumbering
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Northwind traders.docx");

        // Keep source list formatting when importing numbered paragraphs.
        ImportFormatOptions importFormatOptions = new ImportFormatOptions(); { importFormatOptions.setKeepSourceNumbering(true); }
        
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING,
            importFormatOptions);

        ParagraphCollection srcParas = srcDoc.getFirstSection().getBody().getParagraphs();
        for (Paragraph srcPara : (Iterable<Paragraph>) srcParas)
        {
            Node importedNode = importer.importNode(srcPara, false);
            dstDoc.getFirstSection().getBody().appendChild(importedNode);
        }

        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.KeepSourceNumbering.docx");
        //ExEnd:KeepSourceNumbering
    }

    @Test
    public void ignoreTextBoxes() throws Exception
    {
        //ExStart:IgnoreTextBoxes
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Northwind traders.docx");

        // Keep the source text boxes formatting when importing.
        ImportFormatOptions importFormatOptions = new ImportFormatOptions(); { importFormatOptions.setIgnoreTextBoxes(false); }
        
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING,
            importFormatOptions);

        ParagraphCollection srcParas = srcDoc.getFirstSection().getBody().getParagraphs();
        for (Paragraph srcPara : (Iterable<Paragraph>) srcParas)
        {
            Node importedNode = importer.importNode(srcPara, true);
            dstDoc.getFirstSection().getBody().appendChild(importedNode);
        }

        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.IgnoreTextBoxes.docx");
        //ExEnd:IgnoreTextBoxes
    }

    @Test
    public void ignoreHeaderFooter() throws Exception
    {
        //ExStart:IgnoreHeaderFooter
        Document srcDocument = new Document(getMyDir() + "Document source.docx");
        Document dstDocument = new Document(getMyDir() + "Northwind traders.docx");

        ImportFormatOptions importFormatOptions = new ImportFormatOptions(); { importFormatOptions.setIgnoreHeaderFooter(false); }

        dstDocument.appendDocument(srcDocument, ImportFormatMode.KEEP_SOURCE_FORMATTING, importFormatOptions);
        
        dstDocument.save(getArtifactsDir() + "JoinAndAppendDocuments.IgnoreHeaderFooter.docx");
        //ExEnd:IgnoreHeaderFooter
    }

    @Test
    public void linkHeadersFooters() throws Exception
    {
        //ExStart:LinkHeadersFooters
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Northwind traders.docx");

        // Set the appended document to appear on a new page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.NEW_PAGE);
        // Link the headers and footers in the source document to the previous section.
        // This will override any headers or footers already found in the source document.
        srcDoc.getFirstSection().getHeadersFooters().linkToPrevious(true);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.LinkHeadersFooters.docx");
        //ExEnd:LinkHeadersFooters
    }

    @Test
    public void removeSourceHeadersFooters() throws Exception
    {
        //ExStart:RemoveSourceHeadersFooters
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Northwind traders.docx");

        // Remove the headers and footers from each of the sections in the source document.
        for (Section section : (Iterable<Section>) srcDoc.getSections())
        {
            section.clearHeadersFooters();
        }

        // Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting 
        // for HeadersFooters can still be set. This will cause the headers and footers to continue from the destination 
        // document. This should set to false to avoid this behavior.
        srcDoc.getFirstSection().getHeadersFooters().linkToPrevious(false);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.RemoveSourceHeadersFooters.docx");
        //ExEnd:RemoveSourceHeadersFooters
    }

    @Test
    public void unlinkHeadersFooters() throws Exception
    {
        //ExStart:UnlinkHeadersFooters
        Document srcDoc = new Document(getMyDir() + "Document source.docx");
        Document dstDoc = new Document(getMyDir() + "Northwind traders.docx");

        // Unlink the headers and footers in the source document to stop this
        // from continuing the destination document's headers and footers.
        srcDoc.getFirstSection().getHeadersFooters().linkToPrevious(false);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        dstDoc.save(getArtifactsDir() + "JoinAndAppendDocuments.UnlinkHeadersFooters.docx");
        //ExEnd:UnlinkHeadersFooters
    }
}
