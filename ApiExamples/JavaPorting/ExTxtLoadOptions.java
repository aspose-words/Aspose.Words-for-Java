// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.words.TxtLeadingSpacesOptions;
import com.aspose.words.TxtTrailingSpacesOptions;
import com.aspose.words.ParagraphCollection;
import com.aspose.words.DocumentDirection;
import org.testng.annotations.DataProvider;


@Test
public class ExTxtLoadOptions extends ApiExampleBase
{
    @Test (dataProvider = "detectNumberingWithWhitespacesDataProvider")
    public void detectNumberingWithWhitespaces(boolean detectNumberingWithWhitespaces) throws Exception
    {
        //ExStart
        //ExFor:TxtLoadOptions.DetectNumberingWithWhitespaces
        //ExSummary:Shows how lists are detected when plaintext documents are loaded.
        // Create a plaintext document in the form of a string with parts that may be interpreted as lists
        // Upon loading, the first three lists will always be detected by Aspose.Words, and List objects will be created for them after loading
        String textDoc = "Full stop delimiters:\n" +
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

        // The fourth list, with whitespace inbetween the list number and list item contents,
        // will only be detected as a list if "DetectNumberingWithWhitespaces" in a LoadOptions object is set to true,
        // to avoid paragraphs that start with numbers being mistakenly detected as lists
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.setDetectNumberingWithWhitespaces(detectNumberingWithWhitespaces);

        // Load the document while applying LoadOptions as a parameter and verify the result
        Document doc = new Document(new MemoryStream(Encoding.getUTF8().getBytes(textDoc)), loadOptions);

        if (detectNumberingWithWhitespaces)
        {
            Assert.assertEquals(4, doc.getLists().getCount());
            Assert.True(doc.getFirstSection().getBody().getParagraphs().Any(p => p.GetText().Contains("Fourth list") && ((Paragraph)p).IsListItem));
        }
        else
        {
            Assert.assertEquals(3, doc.getLists().getCount());
            Assert.False(doc.getFirstSection().getBody().getParagraphs().Any(p => p.GetText().Contains("Fourth list") && ((Paragraph)p).IsListItem));
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

        TxtLoadOptions loadOptions = new TxtLoadOptions();
        {
            loadOptions.setLeadingSpacesOptions(txtLeadingSpacesOptions);
            loadOptions.setTrailingSpacesOptions(txtTrailingSpacesOptions);
        }

        // Load the document while applying LoadOptions as a parameter and verify the result
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
                Assert.True(paragraphs.All(p => ((Paragraph)p).ParagraphFormat.FirstLineIndent == 0.0d));

                Assert.assertTrue(paragraphs.get(0).getText().startsWith("      Line 1"));
                Assert.assertTrue(paragraphs.get(1).getText().startsWith("    Line 2"));
                Assert.assertTrue(paragraphs.get(2).getText().startsWith(" Line 3"));
                break;
            case TxtLeadingSpacesOptions.TRIM:
                Assert.True(paragraphs.All(p => ((Paragraph)p).ParagraphFormat.FirstLineIndent == 0.0d));

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
        //ExFor:TxtLoadOptions.DocumentDirection
        //ExSummary:Shows how to detect document direction automatically.
        // Create a LoadOptions object and configure it to detect text direction automatically upon loading
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.setDocumentDirection(DocumentDirection.AUTO);
 
        // Text like Hebrew/Arabic will be automatically detected as RTL
        Document doc = new Document(getMyDir() + "Hebrew text.txt", loadOptions);

        Assert.assertTrue(doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getBidi());

        doc = new Document(getMyDir() + "English text.txt", loadOptions);

        Assert.assertFalse(doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getBidi());
        //ExEnd
    }
}
