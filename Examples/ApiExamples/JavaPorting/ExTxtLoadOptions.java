// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.TxtLoadOptions;
import com.aspose.words.Document;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.ms.System.Text.Encoding;
import org.testng.Assert;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.TxtLeadingSpacesOptions;
import com.aspose.words.TxtTrailingSpacesOptions;
import com.aspose.words.ParagraphCollection;
import com.aspose.words.DocumentDirection;
import com.aspose.words.Paragraph;
import com.aspose.words.NodeType;
import com.aspose.ms.System.IO.Stream;
import com.aspose.words.Field;
import com.aspose.ms.System.msConsole;
import org.testng.annotations.DataProvider;


@Test
public class ExTxtLoadOptions extends ApiExampleBase
{
    @Test (dataProvider = "detectNumberingWithWhitespacesDataProvider")
    public void detectNumberingWithWhitespaces(boolean detectNumberingWithWhitespaces) throws Exception
    {
        //ExStart
        //ExFor:TxtLoadOptions.DetectNumberingWithWhitespaces
        //ExSummary:Shows how to detect lists when loading plaintext documents.
        // Create a plaintext document in a string with four separate parts that we may interpret as lists,
        // with different delimiters. Upon loading the plaintext document into a "Document" object,
        // Aspose.Words will always detect the first three lists and will add a "List" object
        // for each to the document's "Lists" property.
        final String TEXT_DOC = "Full stop delimiters:\n" +
                               "1. First list item 1\n" +
                               "2. First list item 2\n" +
                               "3. First list item 3\n\n" +
                               "Right bracket delimiters:\n" +
                               "1) Second list item 1\n" +
                               "2) Second list item 2\n" +
                               "3) Second list item 3\n\n" +
                               "Bullet delimiters:\n" +
                               "• Third list item 1\n" +
                               "• Third list item 2\n" +
                               "• Third list item 3\n\n" +
                               "Whitespace delimiters:\n" +
                               "1 Fourth list item 1\n" +
                               "2 Fourth list item 2\n" +
                               "3 Fourth list item 3";

        // Create a "TxtLoadOptions" object, which we can pass to a document's constructor
        // to modify how we load a plaintext document.
        TxtLoadOptions loadOptions = new TxtLoadOptions();

        // Set the "DetectNumberingWithWhitespaces" property to "true" to detect numbered items
        // with whitespace delimiters, such as the fourth list in our document, as lists.
        // This may also falsely detect paragraphs that begin with numbers as lists.
        // Set the "DetectNumberingWithWhitespaces" property to "false"
        // to not create lists from numbered items with whitespace delimiters.
        loadOptions.setDetectNumberingWithWhitespaces(detectNumberingWithWhitespaces);

        Document doc = new Document(new MemoryStream(Encoding.getUTF8().getBytes(TEXT_DOC)), loadOptions);

        if (detectNumberingWithWhitespaces)
        {
            Assert.assertEquals(4, doc.getLists().getCount());
            Assert.That(doc.getFirstSection().getBody().getParagraphs().Any(p => p.GetText().Contains("Fourth list") && ((Paragraph)p).IsListItem), assertTrue();
        }
        else
        {
            Assert.assertEquals(3, doc.getLists().getCount());
            Assert.That(doc.getFirstSection().getBody().getParagraphs().Any(p => p.GetText().Contains("Fourth list") && ((Paragraph)p).IsListItem), assertFalse();
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "detectNumberingWithWhitespacesDataProvider")
	public static Object[][] detectNumberingWithWhitespacesDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}
    
    @Test (dataProvider = "trailSpacesDataProvider")
    public void trailSpaces(/*TxtLeadingSpacesOptions*/int txtLeadingSpacesOptions, /*TxtTrailingSpacesOptions*/int txtTrailingSpacesOptions) throws Exception
    {
        //ExStart
        //ExFor:TxtLoadOptions.TrailingSpacesOptions
        //ExFor:TxtLoadOptions.LeadingSpacesOptions
        //ExFor:TxtTrailingSpacesOptions
        //ExFor:TxtLeadingSpacesOptions
        //ExSummary:Shows how to trim whitespace when loading plaintext documents.
        String textDoc = "      Line 1 \n" +
                         "    Line 2   \n" +
                         " Line 3       ";

        // Create a "TxtLoadOptions" object, which we can pass to a document's constructor
        // to modify how we load a plaintext document.
        TxtLoadOptions loadOptions = new TxtLoadOptions();

        // Set the "LeadingSpacesOptions" property to "TxtLeadingSpacesOptions.Preserve"
        // to preserve all whitespace characters at the start of every line.
        // Set the "LeadingSpacesOptions" property to "TxtLeadingSpacesOptions.ConvertToIndent"
        // to remove all whitespace characters from the start of every line,
        // and then apply a left first line indent to the paragraph to simulate the effect of the whitespaces.
        // Set the "LeadingSpacesOptions" property to "TxtLeadingSpacesOptions.Trim"
        // to remove all whitespace characters from every line's start.
        loadOptions.setLeadingSpacesOptions(txtLeadingSpacesOptions);

        // Set the "TrailingSpacesOptions" property to "TxtTrailingSpacesOptions.Preserve"
        // to preserve all whitespace characters at the end of every line. 
        // Set the "TrailingSpacesOptions" property to "TxtTrailingSpacesOptions.Trim" to 
        // remove all whitespace characters from the end of every line.
        loadOptions.setTrailingSpacesOptions(txtTrailingSpacesOptions);

        Document doc = new Document(new MemoryStream(Encoding.getUTF8().getBytes(textDoc)), loadOptions);
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        switch (txtLeadingSpacesOptions)
        {
            case TxtLeadingSpacesOptions.CONVERT_TO_INDENT:
                Assert.assertEquals(37.8d, paragraphs.get(0).getParagraphFormat().getFirstLineIndent());
                Assert.assertEquals(25.2d, paragraphs.get(1).getParagraphFormat().getFirstLineIndent());
                Assert.assertEquals(6.3d, paragraphs.get(2).getParagraphFormat().getFirstLineIndent());

                Assert.assertTrue(paragraphs.get(0).getText().startsWith("Line 1"));
                Assert.assertTrue(paragraphs.get(1).getText().startsWith("Line 2"));
                Assert.assertTrue(paragraphs.get(2).getText().startsWith("Line 3"));
                break;
            case TxtLeadingSpacesOptions.PRESERVE:
                Assert.That(paragraphs.All(p => ((Paragraph)p).ParagraphFormat.FirstLineIndent == 0.0d), assertTrue();

                Assert.assertTrue(paragraphs.get(0).getText().startsWith("      Line 1"));
                Assert.assertTrue(paragraphs.get(1).getText().startsWith("    Line 2"));
                Assert.assertTrue(paragraphs.get(2).getText().startsWith(" Line 3"));
                break;
            case TxtLeadingSpacesOptions.TRIM:
                Assert.That(paragraphs.All(p => ((Paragraph)p).ParagraphFormat.FirstLineIndent == 0.0d), assertTrue();

                Assert.assertTrue(paragraphs.get(0).getText().startsWith("Line 1"));
                Assert.assertTrue(paragraphs.get(1).getText().startsWith("Line 2"));
                Assert.assertTrue(paragraphs.get(2).getText().startsWith("Line 3"));
                break;
        }

        switch (txtTrailingSpacesOptions)
        {
            case TxtTrailingSpacesOptions.PRESERVE:
                Assert.assertTrue(paragraphs.get(0).getText().endsWith("Line 1 \r"));
                Assert.assertTrue(paragraphs.get(1).getText().endsWith("Line 2   \r"));
                Assert.assertTrue(paragraphs.get(2).getText().endsWith("Line 3       \f"));
                break;
            case TxtTrailingSpacesOptions.TRIM:
                Assert.assertTrue(paragraphs.get(0).getText().endsWith("Line 1\r"));
                Assert.assertTrue(paragraphs.get(1).getText().endsWith("Line 2\r"));
                Assert.assertTrue(paragraphs.get(2).getText().endsWith("Line 3\f"));
                break;
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "trailSpacesDataProvider")
	public static Object[][] trailSpacesDataProvider() throws Exception
	{
		return new Object[][]
		{
			{TxtLeadingSpacesOptions.PRESERVE,  TxtTrailingSpacesOptions.PRESERVE},
			{TxtLeadingSpacesOptions.CONVERT_TO_INDENT,  TxtTrailingSpacesOptions.PRESERVE},
			{TxtLeadingSpacesOptions.TRIM,  TxtTrailingSpacesOptions.TRIM},
		};
	}

    @Test
    public void detectDocumentDirection() throws Exception
    {
        //ExStart
        //ExFor:DocumentDirection
        //ExFor:TxtLoadOptions.DocumentDirection
        //ExFor:ParagraphFormat.Bidi
        //ExSummary:Shows how to detect plaintext document text direction.
        // Create a "TxtLoadOptions" object, which we can pass to a document's constructor
        // to modify how we load a plaintext document.
        TxtLoadOptions loadOptions = new TxtLoadOptions();

        // Set the "DocumentDirection" property to "DocumentDirection.Auto" automatically detects
        // the direction of every paragraph of text that Aspose.Words loads from plaintext.
        // Each paragraph's "Bidi" property will store its direction.
        loadOptions.setDocumentDirection(DocumentDirection.AUTO);
 
        // Detect Hebrew text as right-to-left.
        Document doc = new Document(getMyDir() + "Hebrew text.txt", loadOptions);

        Assert.assertTrue(doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getBidi());

        // Detect English text as right-to-left.
        doc = new Document(getMyDir() + "English text.txt", loadOptions);

        Assert.assertFalse(doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getBidi());
        //ExEnd
    }

    @Test
    public void autoNumberingDetection() throws Exception
    {
        //ExStart
        //ExFor:TxtLoadOptions.AutoNumberingDetection
        //ExSummary:Shows how to disable automatic numbering detection.
        TxtLoadOptions options = new TxtLoadOptions(); { options.setAutoNumberingDetection(false); }
        Document doc = new Document(getMyDir() + "Number detection.txt", options);
        //ExEnd

        int listItemsCount = 0;
        for (Paragraph paragraph : (Iterable<Paragraph>) doc.getChildNodes(NodeType.PARAGRAPH, true))
        {
            if (paragraph.isListItem())
                listItemsCount++;
        }

        Assert.assertEquals(0, listItemsCount);
    }

    @Test
    public void detectHyperlinks() throws Exception
    {
        //ExStart:DetectHyperlinks
        //GistId:3428e84add5beb0d46a8face6e5fc858
        //ExFor:TxtLoadOptions
        //ExFor:TxtLoadOptions.#ctor
        //ExFor:TxtLoadOptions.DetectHyperlinks
        //ExSummary:Shows how to read and display hyperlinks.
        final String INPUT_TEXT = "Some links in TXT:\n" +
                "https://www.aspose.com/\n" +
                "https://docs.aspose.com/words/net/\n";

        Stream stream = new MemoryStream();
        try /*JAVA: was using*/
        {
            byte[] buf = Encoding.getASCII().getBytes(INPUT_TEXT);
            stream.write(buf, 0, buf.length);

            // Load document with hyperlinks.
            Document doc = new Document(stream, new TxtLoadOptions(); { doc.setDetectHyperlinks(true); });

            // Print hyperlinks text.
            for (Field field : doc.getRange().getFields())
                System.out.println(field.getResult());

            Assert.assertEquals(doc.getRange().getFields().get(0).getResult().trim(), "https://www.aspose.com/");
            Assert.assertEquals(doc.getRange().getFields().get(1).getResult().trim(), "https://docs.aspose.com/words/net/");
        }
        finally { if (stream != null) stream.close(); }
        //ExEnd:DetectHyperlinks
    }
}

