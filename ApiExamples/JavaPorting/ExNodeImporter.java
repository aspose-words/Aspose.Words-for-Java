// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.ImportFormatOptions;
import com.aspose.words.NodeImporter;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.Paragraph;
import com.aspose.words.Node;
import org.testng.Assert;
import com.aspose.words.SaveFormat;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Bookmark;
import com.aspose.words.NodeType;
import com.aspose.words.CompositeNode;
import com.aspose.words.Section;
import com.aspose.words.IFieldMergingCallback;
import com.aspose.words.FieldMergingArgs;
import com.aspose.words.ImageFieldMergingArgs;
import org.testng.annotations.DataProvider;


@Test
public class ExNodeImporter extends ApiExampleBase
{
    @Test (dataProvider = "keepSourceNumberingDataProvider")
    public void keepSourceNumbering(boolean keepSourceNumbering) throws Exception
    {
        //ExStart
        //ExFor:ImportFormatOptions.KeepSourceNumbering
        //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how to resolve list numbering clashes in source and destination documents.
        // Open a document with a custom list numbering scheme, and then clone it.
        // Since both have the same numbering format, the formats will clash if we import one document into the other.
        Document srcDoc = new Document(getMyDir() + "Custom list numbering.docx");
        Document dstDoc = srcDoc.deepClone();

        // When we import the document's clone into the original and then append it,
        // then the two lists with the same list format will join.
        // If we set the "KeepSourceNumbering" flag to "false", then the list from the document clone
        // that we append to the original will carry on the numbering of the list we append it to.
        // This will effectively merge the two lists into one.
        // If we set the "KeepSourceNumbering" flag to "true", then the document clone
        // list will preserve its original numbering, making the two lists appear as separate lists. 
        ImportFormatOptions importFormatOptions = new ImportFormatOptions();
        importFormatOptions.setKeepSourceNumbering(keepSourceNumbering);

        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KEEP_DIFFERENT_STYLES, importFormatOptions);
        for (Paragraph paragraph : (Iterable<Paragraph>) srcDoc.getFirstSection().getBody().getParagraphs())
        {
            Node importedNode = importer.importNode(paragraph, true);
            dstDoc.getFirstSection().getBody().appendChild(importedNode);
        }

        dstDoc.updateListLabels();

        if (keepSourceNumbering)
        {
            Assert.assertEquals(
                "6. Item 1\r\n" +
                "7. Item 2 \r\n" +
                "8. Item 3\r\n" +
                "9. Item 4\r\n" +
                "6. Item 1\r\n" +
                "7. Item 2 \r\n" +
                "8. Item 3\r\n" +
                "9. Item 4", dstDoc.getFirstSection().getBody().toString(SaveFormat.TEXT).trim());
        }
        else
        {
            Assert.assertEquals(
                "6. Item 1\r\n" +
                "7. Item 2 \r\n" +
                "8. Item 3\r\n" +
                "9. Item 4\r\n" +
                "10. Item 1\r\n" +
                "11. Item 2 \r\n" +
                "12. Item 3\r\n" +
                "13. Item 4", dstDoc.getFirstSection().getBody().toString(SaveFormat.TEXT).trim());
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "keepSourceNumberingDataProvider")
	public static Object[][] keepSourceNumberingDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}
    
    //ExStart
    //ExFor:Paragraph.IsEndOfSection
    //ExFor:NodeImporter
    //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode)
    //ExFor:NodeImporter.ImportNode(Node, Boolean)
    //ExSummary:Shows how to insert the contents of one document to a bookmark in another document.
    @Test
    public void insertAtBookmark() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("InsertionPoint");
        builder.write("We will insert a document here: ");
        builder.endBookmark("InsertionPoint");

        Document docToInsert = new Document();
        builder = new DocumentBuilder(docToInsert);

        builder.write("Hello world!");

        docToInsert.save(getArtifactsDir() + "NodeImporter.InsertAtMergeField.docx");

        Bookmark bookmark = doc.getRange().getBookmarks().get("InsertionPoint");
        insertDocument(bookmark.getBookmarkStart().getParentNode(), docToInsert);

        Assert.assertEquals("We will insert a document here: " +
                        "\rHello world!", doc.getText().trim());
    }

    /// <summary>
    /// Inserts the contents of a document after the specified node.
    /// </summary>
    static void insertDocument(Node insertionDestination, Document docToInsert)
    {
        if (insertionDestination.getNodeType() == NodeType.PARAGRAPH || insertionDestination.getNodeType() == NodeType.TABLE)
        {
            CompositeNode destinationParent = insertionDestination.getParentNode();

            NodeImporter importer =
                new NodeImporter(docToInsert, insertionDestination.getDocument(), ImportFormatMode.KEEP_SOURCE_FORMATTING);

            // Loop through all block-level nodes in the section's body,
            // then clone and insert every node that is not the last empty paragraph of a section.
            for (Section srcSection : docToInsert.getSections().<Section>OfType() !!Autoporter error: Undefined expression type )
                for (Node srcNode : (Iterable<Node>) srcSection.getBody())
                {
                    if (srcNode.getNodeType() == NodeType.PARAGRAPH)
                    {
                        Paragraph para = (Paragraph)srcNode;
                        if (para.isEndOfSection() && !para.hasChildNodes())
                            continue;
                    }

                    Node newNode = importer.importNode(srcNode, true);

                    destinationParent.insertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
        }
        else
        {
            throw new IllegalArgumentException("The destination node should be either a paragraph or table.");
        }
    }
    //ExEnd

    @Test
    public void insertAtMergeField() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("A document will appear here: ");
        builder.insertField(" MERGEFIELD Document_1 ");

        Document subDoc = new Document();
        builder = new DocumentBuilder(subDoc);
        builder.write("Hello world!");

        subDoc.save(getArtifactsDir() + "NodeImporter.InsertAtMergeField.docx");

        doc.getMailMerge().setFieldMergingCallback(new InsertDocumentAtMailMergeHandler());

        // The main document has a merge field in it called "Document_1".
        // Execute a mail merge using a data source that contains a local system filename
        // of the document that we wish to insert into the MERGEFIELD.
        doc.getMailMerge().execute(new String[] { "Document_1" },
            new Object[] { getArtifactsDir() + "NodeImporter.InsertAtMergeField.docx" });

        Assert.assertEquals("A document will appear here: \r" +
                        "Hello world!", doc.getText().trim());
    }

    /// <summary>
    /// If the mail merge encounters a MERGEFIELD with a specified name,
    /// this handler treats the current value of a mail merge data source as a local system filename of a document.
    /// The handler will insert the document in its entirety into the MERGEFIELD instead of the current merge value.
    /// </summary>
    private static class InsertDocumentAtMailMergeHandler implements IFieldMergingCallback
    {
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs args) throws Exception
        {
            if ("Document_1".equals(args.getDocumentFieldName()))
            {
                DocumentBuilder builder = new DocumentBuilder(args.getDocument());
                builder.moveToMergeField(args.getDocumentFieldName());

                Document subDoc = new Document((String)args.getFieldValue());

                insertDocument(builder.getCurrentParagraph(), subDoc);

                if (!builder.getCurrentParagraph().hasChildNodes())
                    builder.getCurrentParagraph().remove();

                args.setText(null);
            }
        }

        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs args)
        {
            // Do nothing.
        }
    }
}
