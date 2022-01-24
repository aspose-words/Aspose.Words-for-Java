package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.apache.commons.collections4.IterableUtils;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.ByteArrayInputStream;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

@Test
public class ExTxtLoadOptions extends ApiExampleBase {
    @Test(dataProvider = "detectNumberingWithWhitespacesDataProvider")
    public void detectNumberingWithWhitespaces(boolean detectNumberingWithWhitespaces) throws Exception {
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

        Document doc = new Document(new ByteArrayInputStream(TEXT_DOC.getBytes()), loadOptions);

        List<Paragraph> paragraphList = Arrays.stream(doc.getFirstSection().getBody().getParagraphs().toArray())
                .filter(Paragraph.class::isInstance)
                .map(Paragraph.class::cast)
                .collect(Collectors.toList());

        if (detectNumberingWithWhitespaces) {
            Assert.assertEquals(4, doc.getLists().getCount());
            Assert.assertTrue(IterableUtils.matchesAny(paragraphList, s -> s.getText().contains("Fourth list") && s.isListItem()));
        } else {
            Assert.assertEquals(3, doc.getLists().getCount());
            Assert.assertFalse(IterableUtils.matchesAny(paragraphList, s -> s.getText().contains("Fourth list") && s.isListItem()));
        }
        //ExEnd
    }

    @DataProvider(name = "detectNumberingWithWhitespacesDataProvider")
    public static Object[][] detectNumberingWithWhitespacesDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "trailSpacesDataProvider")
    public void trailSpaces(int txtLeadingSpacesOptions, int txtTrailingSpacesOptions) throws Exception {
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

        Document doc = new Document(new ByteArrayInputStream(textDoc.getBytes()), loadOptions);
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        switch (txtLeadingSpacesOptions) {
            case TxtLeadingSpacesOptions.CONVERT_TO_INDENT:
                Assert.assertEquals(37.8d, paragraphs.get(0).getParagraphFormat().getFirstLineIndent());
                Assert.assertEquals(25.2d, paragraphs.get(1).getParagraphFormat().getFirstLineIndent());
                Assert.assertEquals(6.3d, paragraphs.get(2).getParagraphFormat().getFirstLineIndent());

                Assert.assertTrue(paragraphs.get(0).getText().startsWith("Line 1"));
                Assert.assertTrue(paragraphs.get(1).getText().startsWith("Line 2"));
                Assert.assertTrue(paragraphs.get(2).getText().startsWith("Line 3"));
                break;
            case TxtLeadingSpacesOptions.PRESERVE:
                Assert.assertTrue(IterableUtils.matchesAll(paragraphs, s -> s.getParagraphFormat().getFirstLineIndent() == 0.0d));

                Assert.assertTrue(paragraphs.get(0).getText().startsWith("      Line 1"));
                Assert.assertTrue(paragraphs.get(1).getText().startsWith("    Line 2"));
                Assert.assertTrue(paragraphs.get(2).getText().startsWith(" Line 3"));
                break;
            case TxtLeadingSpacesOptions.TRIM:
                Assert.assertTrue(IterableUtils.matchesAll(paragraphs, s -> s.getParagraphFormat().getFirstLineIndent() == 0.0d));

                Assert.assertTrue(paragraphs.get(0).getText().startsWith("Line 1"));
                Assert.assertTrue(paragraphs.get(1).getText().startsWith("Line 2"));
                Assert.assertTrue(paragraphs.get(2).getText().startsWith("Line 3"));
                break;
        }

        switch (txtTrailingSpacesOptions) {
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

    @DataProvider(name = "trailSpacesDataProvider")
    public static Object[][] trailSpacesDataProvider() {
        return new Object[][]
                {
                        {TxtLeadingSpacesOptions.PRESERVE, TxtTrailingSpacesOptions.PRESERVE},
                        {TxtLeadingSpacesOptions.CONVERT_TO_INDENT, TxtTrailingSpacesOptions.PRESERVE},
                        {TxtLeadingSpacesOptions.TRIM, TxtTrailingSpacesOptions.TRIM},
                };
    }

    @Test
    public void detectDocumentDirection() throws Exception {
        //ExStart
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
}
