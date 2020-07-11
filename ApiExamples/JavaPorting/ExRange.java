// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.ms.System.msConsole;
import com.aspose.words.FindReplaceOptions;
import org.testng.Assert;
import com.aspose.ms.System.msString;
import com.aspose.ms.System.DateTime;
import com.aspose.ms.System.Text.RegularExpressions.Regex;
import com.aspose.words.FieldDocProperty;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.words.FindReplaceDirection;
import com.aspose.words.NodeType;
import com.aspose.words.Run;
import com.aspose.ms.System.Convert;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.ParagraphCollection;
import com.aspose.words.BreakType;
import com.aspose.words.Paragraph;
import com.aspose.words.Node;
import com.aspose.words.CompositeNode;
import com.aspose.words.NodeImporter;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.Section;
import org.testng.annotations.DataProvider;


@Test
public class ExRange extends ApiExampleBase
{
    @Test
    public void replaceSimple() throws Exception
    {
        //ExStart
        //ExFor:Range.Replace(String, String, FindReplaceOptions)
        //ExFor:FindReplaceOptions
        //ExFor:FindReplaceOptions.MatchCase
        //ExFor:FindReplaceOptions.FindWholeWordsOnly
        //ExSummary:Simple find and replace operation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello _CustomerName_,");

        // Check the document contains what we are about to test
        System.out.println(doc.getFirstSection().getBody().getParagraphs().get(0).getText());

        FindReplaceOptions options = new FindReplaceOptions();
        options.setMatchCase(false);
        options.setFindWholeWordsOnly(false);

        doc.getRange().replace("_CustomerName_", "James Bond", options);

        doc.save(getArtifactsDir() + "Range.ReplaceSimple.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Range.ReplaceSimple.docx");

        Assert.assertEquals("Hello James Bond,", msString.trim(doc.getText()));
    }

    @Test (dataProvider = "ignoreDeletedDataProvider")
    public void ignoreDeleted(boolean isIgnoreDeleted) throws Exception
    {
        //ExStart
        //ExFor:FindReplaceOptions.IgnoreDeleted
        //ExSummary:Shows how to ignore text inside delete revisions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
 
        // Insert non-revised text
        builder.writeln("Deleted");
        builder.write("Text");
 
        // Remove first paragraph with tracking revisions
        doc.startTrackRevisionsInternal("John Doe", DateTime.getNow());
        doc.getFirstSection().getBody().getFirstParagraph().remove();
        doc.stopTrackRevisions();
 
        Regex regex = new Regex("e");
        FindReplaceOptions options = new FindReplaceOptions();
 
        // Replace 'e' in document while ignoring/not ignoring deleted text
        options.setIgnoreDeleted(isIgnoreDeleted);
        doc.getRange().replaceInternal(regex, "*", options);

        Assert.assertEquals(msString.trim(doc.getText()), isIgnoreDeleted ? "Deleted\rT*xt" : "D*l*t*d\rT*xt");
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "ignoreDeletedDataProvider")
	public static Object[][] ignoreDeletedDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test (dataProvider = "ignoreInsertedDataProvider")
    public void ignoreInserted(boolean isIgnoreInserted) throws Exception
    {
        //ExStart
        //ExFor:FindReplaceOptions.IgnoreInserted
        //ExSummary:Shows how to ignore text inside insert revisions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
 
        // Insert text with tracking revisions
        doc.startTrackRevisionsInternal("John Doe", DateTime.getNow());
        builder.writeln("Inserted");
        doc.stopTrackRevisions();
 
        // Insert non-revised text
        builder.write("Text");
 
        Regex regex = new Regex("e");
        FindReplaceOptions options = new FindReplaceOptions();
 
        // Replace 'e' in document while ignoring/not ignoring inserted text
        options.setIgnoreInserted(isIgnoreInserted);
        doc.getRange().replaceInternal(regex, "*", options);

        Assert.assertEquals(msString.trim(doc.getText()), isIgnoreInserted ? "Inserted\rT*xt" : "Ins*rt*d\rT*xt");
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "ignoreInsertedDataProvider")
	public static Object[][] ignoreInsertedDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test (dataProvider = "ignoreFieldsDataProvider")
    public void ignoreFields(boolean isIgnoreFields) throws Exception
    {
        //ExStart
        //ExFor:FindReplaceOptions.IgnoreFields
        //ExSummary:Shows how to ignore text inside fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
 
        // Insert field with text inside
        builder.insertField("INCLUDETEXT", "Text in field");
 
        Regex regex = new Regex("e");
        FindReplaceOptions options = new FindReplaceOptions();
        // Replace 'e' in document ignoring/not ignoring text inside field
        options.setIgnoreFields(isIgnoreFields);
        
        doc.getRange().replaceInternal(regex, "*", options);

        Assert.assertEquals(doc.getText(),
            isIgnoreFields
                ? "\u0013INCLUDETEXT\u0014Text in field\u0015\f"
                : "\u0013INCLUDETEXT\u0014T*xt in fi*ld\u0015\f");
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "ignoreFieldsDataProvider")
	public static Object[][] ignoreFieldsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test
    public void updateFieldsInRange() throws Exception
    {
        //ExStart
        //ExFor:Range.UpdateFields
        //ExSummary:Shows how to update document fields in the body of the first section only.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a field that will display the value in the document's body text
        FieldDocProperty field = (FieldDocProperty)builder.insertField(" DOCPROPERTY Category");

        // Set the value of the property that should be displayed by the field
        doc.getBuiltInDocumentProperties().setCategory("MyCategory");

        // Some field types need to be explicitly updated before they can display their expected values
        Assert.assertEquals("", field.getResult());

        // Update all the fields in the first section of the document, which includes the field we just inserted
        doc.getFirstSection().getRange().updateFields();

        Assert.assertEquals("MyCategory", field.getResult());
        //ExEnd
    }

    @Test
    public void replaceWithString() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("This one is sad.");
        builder.writeln("That one is mad.");

        FindReplaceOptions options = new FindReplaceOptions();
        options.setMatchCase(false);
        options.setFindWholeWordsOnly(true);

        doc.getRange().replace("sad", "bad", options);

        doc.save(getArtifactsDir() + "Range.ReplaceWithString.docx");
    }

    @Test
    public void replaceWithRegex() throws Exception
    {
        //ExStart
        //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
        //ExSummary:Shows how to replace all occurrences of words "sad" or "mad" to "bad".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("sad mad bad");

        Assert.assertEquals("sad mad bad", msString.trim(doc.getText()));

        FindReplaceOptions options = new FindReplaceOptions();
        {
            options.setMatchCase(false);
            options.setFindWholeWordsOnly(false);
        }

        doc.getRange().replaceInternal(new Regex("[s|m]ad"), "bad", options);

        Assert.assertEquals("bad bad bad", msString.trim(doc.getText()));
        //ExEnd
    }

    //ExStart
    //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
    //ExFor:ReplacingArgs.Replacement
    //ExFor:IReplacingCallback
    //ExFor:IReplacingCallback.Replacing
    //ExFor:ReplacingArgs
    //ExSummary:Replaces text specified with regular expression with HTML.
    @Test //ExSkip
    public void replaceWithInsertHtml() throws Exception
    {
        // Open the document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello <CustomerName>,");

        FindReplaceOptions options = new FindReplaceOptions();
        options.setReplacingCallback(new ReplaceWithHtmlEvaluator());

        doc.getRange().replaceInternal(new Regex(" <CustomerName>,"), "", options);

        // Save the modified document
        doc.save(getArtifactsDir() + "Range.ReplaceWithInsertHtml.docx");
        Assert.assertEquals("James Bond, Hello\r\f",
            new Document(getArtifactsDir() + "Range.ReplaceWithInsertHtml.docx").getText()); //ExSkip
    }

    private static class ReplaceWithHtmlEvaluator implements IReplacingCallback
    {
        public /*ReplaceAction*/int /*IReplacingCallback.*/replacing(ReplacingArgs args) throws Exception
        {
            DocumentBuilder builder = new DocumentBuilder((Document)args.getMatchNode().getDocument());
            builder.moveTo(args.getMatchNode());

            // Replace '<CustomerName>' text with a red bold name
            builder.insertHtml("<b><font color='red'>James Bond, </font></b>");
            args.setReplacement("");

            return ReplaceAction.REPLACE;
        }
    }
    //ExEnd

    //ExStart
    //ExFor:FindReplaceOptions.ApplyFont
    //ExFor:FindReplaceOptions.Direction
    //ExFor:FindReplaceOptions.ReplacingCallback
    //ExFor:ReplacingArgs.GroupIndex
    //ExFor:ReplacingArgs.GroupName
    //ExFor:ReplacingArgs.Match
    //ExFor:ReplacingArgs.MatchOffset
    //ExSummary:Shows how to apply a different font to new content via FindReplaceOptions.
    @Test //ExSkip
    public void replaceNumbersAsHex() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Arial");
        builder.writeln("Numbers that will be converted to hexadecimal and highlighted:\n" +
                        "123, 456, 789 and 17379.");

        FindReplaceOptions options = new FindReplaceOptions();
        // Highlight newly inserted content with a color
        options.getApplyFont().setHighlightColor(msColor.getLightGray());
        // Apply an IReplacingCallback to make the replacement to convert integers into hex equivalents
        // and also to count replacements in the order they take place
        options.setReplacingCallback(new NumberHexer());
        // By default, text is searched for replacements front to back, but we can change it to go the other way
        options.setDirection(FindReplaceDirection.BACKWARD);

        int count = doc.getRange().replaceInternal(new Regex("[0-9]+"), "", options);

        Assert.assertEquals(4, count);
        Assert.assertEquals("Numbers that will be converted to hexadecimal and highlighted:\r" +
                        "0x7B, 0x1C8, 0x315 and 0x43E3.", msString.trim(doc.getText()));
        Assert.AreEqual(4,
            doc.getChildNodes(NodeType.RUN, true).<Run>OfType()
                .Count(r => r.Font.HighlightColor.ToArgb() == Color.LightGray.ToArgb()));
    }

    /// <summary>
    /// Replaces arabic numbers with hexadecimal equivalents and appends the number of each replacement.
    /// </summary>
    private static class NumberHexer implements IReplacingCallback
    {
        public /*ReplaceAction*/int replacing(ReplacingArgs args)
        {
            mCurrentReplacementNumber++;
            
            // Parse numbers
            int number = Convert.toInt32(args.getMatchInternal().getValue());

            // And write it as HEX
            args.setReplacement("0x{number:X}");

            System.out.println("Match #{mCurrentReplacementNumber}");
            System.out.println("\tOriginal value:\t{args.Match.Value}");
            System.out.println("\tReplacement:\t{args.Replacement}");
            System.out.println("\tOffset in parent {args.MatchNode.NodeType} node:\t{args.MatchOffset}");

            System.out.println(msString.isNullOrEmpty(args.GroupName)
                    ? $"\tGroup index:\t{args.GroupIndex}"
                    : $"\tGroup name:\t{args.GroupName}");

            return ReplaceAction.REPLACE;
        }

        private int mCurrentReplacementNumber;
    }
    //ExEnd

    @Test
    public void applyParagraphFormat() throws Exception
    {
        //ExStart
        //ExFor:FindReplaceOptions.ApplyParagraphFormat
        //ExSummary:Shows how to affect the format of paragraphs with successful replacements.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Every paragraph that ends with a full stop like this one will be right aligned.");
        builder.writeln("This one will not!");
        builder.writeln("And this one will.");
        
        FindReplaceOptions options = new FindReplaceOptions();
        options.getApplyParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);

        int count = doc.getRange().replace(".&p", "!&p", options);
        Assert.assertEquals(2, count);

        doc.save(getArtifactsDir() + "Range.ApplyParagraphFormat.docx");
        //ExEnd

        ParagraphCollection paragraphs = new Document(getArtifactsDir() + "Range.ApplyParagraphFormat.docx").getFirstSection().getBody().getParagraphs();

        Assert.assertEquals(ParagraphAlignment.RIGHT, paragraphs.get(0).getParagraphFormat().getAlignment());
        Assert.assertEquals(ParagraphAlignment.LEFT, paragraphs.get(1).getParagraphFormat().getAlignment());
        Assert.assertEquals(ParagraphAlignment.RIGHT, paragraphs.get(2).getParagraphFormat().getAlignment());
    }

    @Test
    public void deleteSelection() throws Exception
    {
        //ExStart
        //ExFor:Node.Range
        //ExFor:Range.Delete
        //ExSummary:Shows how to delete all characters of a range.
        // Insert two sections into a blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Section 1. ");
        builder.insertBreak(BreakType.SECTION_BREAK_CONTINUOUS);
        builder.write("Section 2.");

        // Verify the whole text of the document
        Assert.assertEquals("Section 1. \fSection 2.", msString.trim(doc.getText()));

        // Delete the first section from the document
        doc.getSections().get(0).getRange().delete();

        // Check the first section was deleted by looking at the text of the whole document again
        Assert.assertEquals("Section 2.", msString.trim(doc.getText()));
        //ExEnd
    }

    @Test
    public void rangesGetText() throws Exception
    {
        //ExStart
        //ExFor:Range
        //ExFor:Range.Text
        //ExSummary:Shows how to get plain, unformatted text of a range.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Hello world!");

        Assert.assertEquals("Hello world!", msString.trim(doc.getRange().getText()));
        //ExEnd
    }

    //ExStart
    //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
    //ExFor:IReplacingCallback
    //ExFor:ReplaceAction
    //ExFor:IReplacingCallback.Replacing
    //ExFor:ReplacingArgs
    //ExFor:ReplacingArgs.MatchNode
    //ExFor:FindReplaceDirection
    //ExSummary:Shows how to insert content of one document into another during a customized find and replace operation.
    @Test //ExSkip
    public void insertDocumentAtReplace() throws Exception
    {
        Document mainDoc = new Document(getMyDir() + "Document insertion destination.docx");

        FindReplaceOptions options = new FindReplaceOptions();
        options.setDirection(FindReplaceDirection.BACKWARD);
        options.setReplacingCallback(new InsertDocumentAtReplaceHandler());

        mainDoc.getRange().replaceInternal(new Regex("\\[MY_DOCUMENT\\]"), "", options);
        mainDoc.save(getArtifactsDir() + "InsertDocument.InsertDocumentAtReplace.docx");

        testInsertDocumentAtReplace(new Document(getArtifactsDir() + "InsertDocument.InsertDocumentAtReplace.docx")); //ExSkip
    }

    private static class InsertDocumentAtReplaceHandler implements IReplacingCallback
    {
        public /*ReplaceAction*/int /*IReplacingCallback.*/replacing(ReplacingArgs args) throws Exception
        {
            Document subDoc = new Document(getMyDir() + "Document.docx");

            // Insert a document after the paragraph, containing the match text
            Paragraph para = (Paragraph)args.getMatchNode().getParentNode();
            insertDocument(para, subDoc);

            // Remove the paragraph with the match text
            para.remove();

            return ReplaceAction.SKIP;
        }
    }

    /// <summary>
    /// Inserts content of the external document after the specified node.
    /// </summary>
    static void insertDocument(Node insertionDestination, Document docToInsert)
    {
        // Make sure that the node is either a paragraph or table
        if (((insertionDestination.getNodeType()) == (NodeType.PARAGRAPH)) || ((insertionDestination.getNodeType()) == (NodeType.TABLE)))
        {
            // We will be inserting into the parent of the destination paragraph
            CompositeNode dstStory = insertionDestination.getParentNode();

            // This object will be translating styles and lists during the import
            NodeImporter importer =
                new NodeImporter(docToInsert, insertionDestination.getDocument(), ImportFormatMode.KEEP_SOURCE_FORMATTING);

            // Loop through all block level nodes in the body of the section
            for (Section srcSection : docToInsert.getSections().<Section>OfType() !!Autoporter error: Undefined expression type )
                for (Node srcNode : (Iterable<Node>) srcSection.getBody())
                {
                    // Let's skip the node if it is a last empty paragraph in a section
                    if (((srcNode.getNodeType()) == (NodeType.PARAGRAPH)))
                    {
                        Paragraph para = (Paragraph)srcNode;
                        if (para.isEndOfSection() && !para.hasChildNodes())
                            continue;
                    }

                    // This creates a clone of the node, suitable for insertion into the destination document
                    Node newNode = importer.importNode(srcNode, true);

                    // Insert new node after the reference node
                    dstStory.insertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
        }
        else
        {
            throw new IllegalArgumentException("The destination node should be either a paragraph or table.");
        }
    }
    //ExEnd

    private void testInsertDocumentAtReplace(Document doc)
    {
        Assert.assertEquals("1) At text that can be identified by regex:\rHello World!\r" +
                        "2) At a MERGEFIELD:\r\u0013 MERGEFIELD  Document_1  \\* MERGEFORMAT \u0014«Document_1»\u0015\r" +
                        "3) At a bookmark:", msString.trim(doc.getFirstSection().getBody().getText()));
    }
}
