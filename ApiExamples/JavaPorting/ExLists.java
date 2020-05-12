// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.ms;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import org.testng.Assert;
import com.aspose.words.NumberStyle;
import com.aspose.words.ListTemplate;
import com.aspose.words.List;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.ListLevel;
import java.awt.Color;
import com.aspose.words.ListLevelAlignment;
import com.aspose.words.ListTrailingCharacter;
import com.aspose.words.Style;
import com.aspose.words.StyleType;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.ms.System.msConsole;
import com.aspose.words.BreakType;
import com.aspose.words.ListCollection;
import com.aspose.ms.System.msString;
import com.aspose.words.SaveFormat;
import com.aspose.words.ListLabel;


@Test
public class ExLists extends ApiExampleBase
{
    @Test
    public void applyDefaultBulletsAndNumbers() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.ListFormat
        //ExFor:ListFormat.ApplyNumberDefault
        //ExFor:ListFormat.ApplyBulletDefault
        //ExFor:ListFormat.ListIndent
        //ExFor:ListFormat.ListOutdent
        //ExFor:ListFormat.RemoveNumbers
        //ExFor:ListFormat.ListLevelNumber
        //ExSummary:Shows how to apply default bulleted or numbered list formatting to paragraphs when using DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Aspose.Words allows:");
        builder.writeln();

        // Start a numbered list with default formatting
        builder.getListFormat().applyNumberDefault();
        builder.writeln("Opening documents from different formats:");

        Assert.assertEquals(0, builder.getListFormat().getListLevelNumber());

        // Go to second list level, add more text
        builder.getListFormat().listIndent();

        Assert.assertEquals(1, builder.getListFormat().getListLevelNumber());

        builder.writeln("DOC");
        builder.writeln("PDF");
        builder.writeln("HTML");

        // Outdent to the first list level
        builder.getListFormat().listOutdent();

        Assert.assertEquals(0, builder.getListFormat().getListLevelNumber());

        builder.writeln("Processing documents");
        builder.writeln("Saving documents in different formats:");

        // Indent the list level again
        builder.getListFormat().listIndent();
        builder.writeln("DOC");
        builder.writeln("PDF");
        builder.writeln("HTML");
        builder.writeln("MHTML");
        builder.writeln("Plain text");

        // Outdent the list level again
        builder.getListFormat().listOutdent();
        builder.writeln("Doing many other things!");

        // End the numbered list
        builder.getListFormat().removeNumbers();
        builder.writeln();

        builder.writeln("Aspose.Words main advantages are:");
        builder.writeln();

        // Start a bulleted list with default formatting
        builder.getListFormat().applyBulletDefault();
        builder.writeln("Great performance");
        builder.writeln("High reliability");
        builder.writeln("Quality code and working");
        builder.writeln("Wide variety of features");
        builder.writeln("Easy to understand API");

        // End the bulleted list
        builder.getListFormat().removeNumbers();

        doc.save(getArtifactsDir() + "Lists.ApplyDefaultBulletsAndNumbers.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Lists.ApplyDefaultBulletsAndNumbers.docx");

        TestUtil.verifyListLevel("\u0000.", 18.0d, NumberStyle.ARABIC, doc.getLists().get(0).getListLevels().get(0));
        TestUtil.verifyListLevel("\u0001.", 54.0d, NumberStyle.LOWERCASE_LETTER, doc.getLists().get(0).getListLevels().get(1));
        TestUtil.verifyListLevel("\uf0b7", 18.0d, NumberStyle.BULLET, doc.getLists().get(1).getListLevels().get(0));
    }

    @Test
    public void specifyListLevel() throws Exception
    {
        //ExStart
        //ExFor:ListCollection
        //ExFor:List
        //ExFor:ListFormat
        //ExFor:ListFormat.ListLevelNumber
        //ExFor:ListFormat.List
        //ExFor:ListTemplate
        //ExFor:DocumentBase.Lists
        //ExFor:ListCollection.Add(ListTemplate)
        //ExSummary:Shows how to specify list level number when building a list using DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list based on one of the Microsoft Word list templates and
        // apply it to the current paragraph in the document builder
        builder.getListFormat().setList(doc.getLists().add(ListTemplate.NUMBER_ARABIC_DOT));

        // Insert text at each of the 9 indent levels
        for (int i = 0; i < 9; i++)
        {
            builder.getListFormat().setListLevelNumber(i);
            builder.writeln("Level " + i);
        }

        // Create a bulleted list based on one of the Microsoft Word list templates
        // and apply it to the current paragraph in the document builder
        builder.getListFormat().setList(doc.getLists().add(ListTemplate.BULLET_DIAMONDS));

        for (int i = 0; i < 9; i++)
        {
            builder.getListFormat().setListLevelNumber(i);
            builder.writeln("Level " + i);
        }

        // This is a way to stop list formatting
        builder.getListFormat().setList(null);

        doc.save(getArtifactsDir() + "Lists.SpecifyListLevel.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Lists.SpecifyListLevel.docx");

        TestUtil.verifyListLevel("\u0000.", 18.0d, NumberStyle.ARABIC, doc.getLists().get(0).getListLevels().get(0));
    }

    @Test
    public void nestedLists() throws Exception
    {
        //ExStart
        //ExFor:ListFormat.List
        //ExFor:ParagraphFormat.ClearFormatting
        //ExFor:ParagraphFormat.DropCapPosition
        //ExFor:ParagraphFormat.IsListItem
        //ExFor:Paragraph.IsListItem
        //ExSummary:Shows how to start a numbered list, add a bulleted list inside it, then return to the numbered list.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an outline list for the headings
        List outlineList = doc.getLists().add(ListTemplate.OUTLINE_NUMBERS);
        builder.getListFormat().setList(outlineList);
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.writeln("This is my Chapter 1");

        // Create a numbered list
        List numberedList = doc.getLists().add(ListTemplate.NUMBER_DEFAULT);
        builder.getListFormat().setList(numberedList);
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.NORMAL);
        builder.writeln("Numbered list item 1.");

        // Every paragraph that comprises a list will have this flag
        Assert.assertTrue(builder.getCurrentParagraph().isListItem());
        Assert.assertTrue(builder.getParagraphFormat().isListItem());

        // Create a bulleted list
        List bulletedList = doc.getLists().add(ListTemplate.BULLET_DEFAULT);
        builder.getListFormat().setList(bulletedList);
        builder.getParagraphFormat().setLeftIndent(72.0);
        builder.writeln("Bulleted list item 1.");
        builder.writeln("Bulleted list item 2.");
        builder.getParagraphFormat().clearFormatting();

        // Revert to the numbered list
        builder.getListFormat().setList(numberedList);
        builder.writeln("Numbered list item 2.");
        builder.writeln("Numbered list item 3.");

        // Revert to the outline list
        builder.getListFormat().setList(outlineList);
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.writeln("This is my Chapter 2");

        builder.getParagraphFormat().clearFormatting();

        builder.getDocument().save(getArtifactsDir() + "Lists.NestedLists.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Lists.NestedLists.docx");

        TestUtil.verifyListLevel("\u0000)", 0.0d, NumberStyle.ARABIC, doc.getLists().get(0).getListLevels().get(0));
        TestUtil.verifyListLevel("\u0000.", 18.0d, NumberStyle.ARABIC, doc.getLists().get(1).getListLevels().get(0));
        TestUtil.verifyListLevel("\uf0b7", 18.0d, NumberStyle.BULLET, doc.getLists().get(2).getListLevels().get(0));
    }

    @Test
    public void createCustomList() throws Exception
    {
        //ExStart
        //ExFor:List
        //ExFor:List.ListLevels
        //ExFor:ListFormat.ListLevel
        //ExFor:ListLevelCollection
        //ExFor:ListLevelCollection.Item
        //ExFor:ListLevel
        //ExFor:ListLevel.Alignment
        //ExFor:ListLevel.Font
        //ExFor:ListLevel.NumberStyle
        //ExFor:ListLevel.StartAt
        //ExFor:ListLevel.TrailingCharacter
        //ExFor:ListLevelAlignment
        //ExFor:NumberStyle
        //ExFor:ListTrailingCharacter
        //ExFor:ListLevel.NumberFormat
        //ExFor:ListLevel.NumberPosition
        //ExFor:ListLevel.TextPosition
        //ExFor:ListLevel.TabPosition
        //ExSummary:Shows how to apply custom list formatting to paragraphs when using DocumentBuilder.
        Document doc = new Document();

        // Create a list based on one of the Microsoft Word list templates
        List list = doc.getLists().add(ListTemplate.NUMBER_DEFAULT);

        // Completely customize one list level
        ListLevel listLevel = list.getListLevels().get(0);
        listLevel.getFont().setColor(Color.RED);
        listLevel.getFont().setSize(24.0);
        listLevel.setNumberStyle(NumberStyle.ORDINAL_TEXT);
        listLevel.setStartAt(21);
        listLevel.setNumberFormat("\u0000");

        listLevel.setNumberPosition(-36);
        listLevel.setTextPosition(144.0);
        listLevel.setTabPosition(144.0);

        // Customize another list level
        listLevel = list.getListLevels().get(1);
        listLevel.setAlignment(ListLevelAlignment.RIGHT);
        listLevel.setNumberStyle(NumberStyle.BULLET);
        listLevel.getFont().setName("Wingdings");
        listLevel.getFont().setColor(Color.BLUE);
        listLevel.getFont().setSize(24.0);
        listLevel.setNumberFormat("\uf0af"); // A bullet that looks like a star
        listLevel.setTrailingCharacter(ListTrailingCharacter.SPACE);
        listLevel.setNumberPosition(144.0);

        // Now add some text that uses the list that we created
        // It does not matter when to customize the list - before or after adding the paragraphs
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getListFormat().setList(list);
        builder.writeln("The quick brown fox...");
        builder.writeln("The quick brown fox...");

        builder.getListFormat().listIndent();
        builder.writeln("jumped over the lazy dog.");
        builder.writeln("jumped over the lazy dog.");

        builder.getListFormat().listOutdent();
        builder.writeln("The quick brown fox...");

        builder.getListFormat().removeNumbers();

        builder.getDocument().save(getArtifactsDir() + "Lists.CreateCustomList.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Lists.CreateCustomList.docx");

        listLevel = doc.getLists().get(0).getListLevels().get(0);

        TestUtil.verifyListLevel("\u0000", -36.0d, NumberStyle.ORDINAL_TEXT, listLevel);
        Assert.assertEquals(Color.RED.getRGB(), listLevel.getFont().getColor().getRGB());
        Assert.assertEquals(24.0d, listLevel.getFont().getSize());
        Assert.assertEquals(21, listLevel.getStartAt());

        listLevel = doc.getLists().get(0).getListLevels().get(1);

        TestUtil.verifyListLevel("\uf0af", 144.0d, NumberStyle.BULLET, listLevel);
        Assert.assertEquals(Color.BLUE.getRGB(), listLevel.getFont().getColor().getRGB());
        Assert.assertEquals(24.0d, listLevel.getFont().getSize());
        Assert.assertEquals(1, listLevel.getStartAt());
        Assert.assertEquals(ListTrailingCharacter.SPACE, listLevel.getTrailingCharacter());
    }

    @Test
    public void restartNumberingUsingListCopy() throws Exception
    {
        //ExStart
        //ExFor:List
        //ExFor:ListCollection
        //ExFor:ListCollection.Add(ListTemplate)
        //ExFor:ListCollection.AddCopy(List)
        //ExFor:ListLevel.StartAt
        //ExFor:ListTemplate
        //ExSummary:Shows how to restart numbering in a list by copying a list.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list based on a template
        List list1 = doc.getLists().add(ListTemplate.NUMBER_ARABIC_PARENTHESIS);
        // Modify the formatting of the list
        list1.getListLevels().get(0).getFont().setColor(Color.RED);
        list1.getListLevels().get(0).setAlignment(ListLevelAlignment.RIGHT);

        builder.writeln("List 1 starts below:");
        // Use the first list in the document for a while
        builder.getListFormat().setList(list1);
        builder.writeln("Item 1");
        builder.writeln("Item 2");
        builder.getListFormat().removeNumbers();

        // Now I want to reuse the first list, but need to restart numbering
        // This should be done by creating a copy of the original list formatting
        List list2 = doc.getLists().addCopy(list1);

        // We can modify the new list in any way. Including setting new start number
        list2.getListLevels().get(0).setStartAt(10);

        // Use the second list in the document
        builder.writeln("List 2 starts below:");
        builder.getListFormat().setList(list2);
        builder.writeln("Item 1");
        builder.writeln("Item 2");
        builder.getListFormat().removeNumbers();

        doc.save(getArtifactsDir() + "Lists.RestartNumberingUsingListCopy.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Lists.RestartNumberingUsingListCopy.docx");

        list1 = doc.getLists().get(0);
        TestUtil.verifyListLevel("\u0000)", 18.0d, NumberStyle.ARABIC, list1.getListLevels().get(0));
        Assert.assertEquals(Color.RED.getRGB(), list1.getListLevels().get(0).getFont().getColor().getRGB());
        Assert.assertEquals(10.0d, list1.getListLevels().get(0).getFont().getSize());
        Assert.assertEquals(1, list1.getListLevels().get(0).getStartAt());

        list2 = doc.getLists().get(1);
        TestUtil.verifyListLevel("\u0000)", 18.0d, NumberStyle.ARABIC, list2.getListLevels().get(0));
        Assert.assertEquals(Color.RED.getRGB(), list2.getListLevels().get(0).getFont().getColor().getRGB());
        Assert.assertEquals(10.0d, list2.getListLevels().get(0).getFont().getSize());
        Assert.assertEquals(10, list2.getListLevels().get(0).getStartAt());
    }

    @Test
    public void createAndUseListStyle() throws Exception
    {
        //ExStart
        //ExFor:StyleCollection.Add(StyleType,String)
        //ExFor:Style.List
        //ExFor:StyleType
        //ExFor:List.IsListStyleDefinition
        //ExFor:List.IsListStyleReference
        //ExFor:List.IsMultiLevel
        //ExFor:List.Style
        //ExFor:ListLevelCollection
        //ExFor:ListLevelCollection.Count
        //ExFor:ListLevelCollection.Item
        //ExFor:ListCollection.Add(Style)
        //ExSummary:Shows how to create a list style and use it in a document.
        Document doc = new Document();

        // Create a new list style
        // List formatting associated with this list style is default numbered
        Style listStyle = doc.getStyles().add(StyleType.LIST, "MyListStyle");

        // This list defines the formatting of the list style
        // Note this list can not be used directly to apply formatting to paragraphs (see below)
        List list1 = listStyle.getList();

        // Check some basic rules about the list that defines a list style
        Assert.assertTrue(list1.isListStyleDefinition());
        Assert.assertFalse(list1.isListStyleReference());
        Assert.assertTrue(list1.isMultiLevel());
        Assert.assertEquals(listStyle, list1.getStyle());

        // Modify formatting of the list style to our liking
        for (ListLevel level : list1.getListLevels())
        {
            level.getFont().setName("Verdana");
            level.getFont().setColor(Color.BLUE);
            level.getFont().setBold(true);
        }

        // Add some text to our document and use the list style
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Using list style first time:");

        // This creates a list based on the list style
        List list2 = doc.getLists().add(listStyle);

        // Check some basic rules about the list that references a list style
        Assert.assertFalse(list2.isListStyleDefinition());
        Assert.assertTrue(list2.isListStyleReference());
        Assert.assertEquals(listStyle, list2.getStyle());

        // Apply the list that references the list style
        builder.getListFormat().setList(list2);
        builder.writeln("Item 1");
        builder.writeln("Item 2");
        builder.getListFormat().removeNumbers();

        builder.writeln("Using list style second time:");

        // Create and apply another list based on the list style
        List list3 = doc.getLists().add(listStyle);
        builder.getListFormat().setList(list3);
        builder.writeln("Item 1");
        builder.writeln("Item 2");
        builder.getListFormat().removeNumbers();

        builder.getDocument().save(getArtifactsDir() + "Lists.CreateAndUseListStyle.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Lists.CreateAndUseListStyle.docx");

        list1 = doc.getLists().get(0);

        TestUtil.verifyListLevel("\u0000.", 18.0d, NumberStyle.ARABIC, list1.getListLevels().get(0));
        Assert.assertTrue(list1.isListStyleDefinition());
        Assert.assertFalse(list1.isListStyleReference());
        Assert.assertTrue(list1.isMultiLevel());
        Assert.assertEquals(Color.BLUE.getRGB(), list1.getListLevels().get(0).getFont().getColor().getRGB());
        Assert.assertEquals("Verdana", list1.getListLevels().get(0).getFont().getName());
        Assert.assertTrue(list1.getListLevels().get(0).getFont().getBold());

        list2 = doc.getLists().get(1);

        TestUtil.verifyListLevel("\u0000.", 18.0d, NumberStyle.ARABIC, list2.getListLevels().get(0));
        Assert.assertFalse(list2.isListStyleDefinition());
        Assert.assertTrue(list2.isListStyleReference());
        Assert.assertTrue(list2.isMultiLevel());

        list3 = doc.getLists().get(2);

        TestUtil.verifyListLevel("\u0000.", 18.0d, NumberStyle.ARABIC, list3.getListLevels().get(0));
        Assert.assertFalse(list3.isListStyleDefinition());
        Assert.assertTrue(list3.isListStyleReference());
        Assert.assertTrue(list3.isMultiLevel());
    }

    @Test
    public void detectBulletedParagraphs() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.ListFormat
        //ExFor:ListFormat.IsListItem
        //ExFor:CompositeNode.GetText
        //ExFor:List.ListId
        //ExSummary:Shows how to output all paragraphs in a document that are bulleted or numbered.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getListFormat().applyNumberDefault();
        builder.writeln("Numbered list item 1");
        builder.writeln("Numbered list item 2");
        builder.writeln("Numbered list item 3");
        builder.getListFormat().removeNumbers();

        builder.getListFormat().applyBulletDefault();
        builder.writeln("Bulleted list item 1");
        builder.writeln("Bulleted list item 2");
        builder.writeln("Bulleted list item 3");
        builder.getListFormat().removeNumbers();

        NodeCollection paras = doc.getChildNodes(NodeType.PARAGRAPH, true);

        for (Paragraph para : paras.<Paragraph>OfType().Where(p => p.ListFormat.IsListItem) !!Autoporter error: Undefined expression type )
        { 
            System.out.println("This paragraph belongs to list ID# {para.ListFormat.List.ListId}, number style \"{para.ListFormat.ListLevel.NumberStyle}\"");
            System.out.println("\t\"{para.GetText().Trim()}\"");
        }
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        paras = doc.getChildNodes(NodeType.PARAGRAPH, true);

        Assert.AreEqual(6, paras.Count(n => (ms.as(n, Paragraph.class)).ListFormat.IsListItem));
    }

    @Test
    public void removeBulletsFromParagraphs() throws Exception
    {
        //ExStart
        //ExFor:ListFormat.RemoveNumbers
        //ExSummary:Shows how to remove bullets and numbering from all paragraphs in the main text of a section.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getListFormat().applyNumberDefault();
        builder.writeln("Numbered list item 1");
        builder.writeln("Numbered list item 2");
        builder.writeln("Numbered list item 3");
        builder.getListFormat().removeNumbers();

        NodeCollection paras = doc.getChildNodes(NodeType.PARAGRAPH, true);

        Assert.AreEqual(3, paras.Count(n => (ms.as(n, Paragraph.class)).ListFormat.IsListItem));

        for (Paragraph paragraph : (Iterable<Paragraph>) paras)
            paragraph.getListFormat().removeNumbers();

        Assert.AreEqual(0, paras.Count(n => (ms.as(n, Paragraph.class)).ListFormat.IsListItem));
        //ExEnd
    }

    @Test
    public void applyExistingListToParagraphs() throws Exception
    {
        //ExStart
        //ExFor:ListCollection.Item(Int32)
        //ExSummary:Shows how to apply list formatting of an existing list to a collection of paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Paragraph 1");
        builder.writeln("Paragraph 2");
        builder.write("Paragraph 3");

        NodeCollection paras = doc.getChildNodes(NodeType.PARAGRAPH, true);

        Assert.AreEqual(0, paras.Count(n => (ms.as(n, Paragraph.class)).ListFormat.IsListItem));

        doc.getLists().add(ListTemplate.NUMBER_DEFAULT);
        List list = doc.getLists().get(0);

        for (Paragraph paragraph : paras.<Paragraph>OfType() !!Autoporter error: Undefined expression type )
        {
            paragraph.getListFormat().setList(list);
            paragraph.getListFormat().setListLevelNumber(2);
        }

        Assert.AreEqual(3, paras.Count(n => (ms.as(n, Paragraph.class)).ListFormat.IsListItem));
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        paras = doc.getChildNodes(NodeType.PARAGRAPH, true);

        Assert.AreEqual(3, paras.Count(n => (ms.as(n, Paragraph.class)).ListFormat.IsListItem));
        Assert.AreEqual(3, paras.Count(n => (ms.as(n, Paragraph.class)).ListFormat.ListLevelNumber == 2));
    }

    @Test
    public void applyNewListToParagraphs() throws Exception
    {
        //ExStart
        //ExFor:ListCollection.Add(ListTemplate)
        //ExSummary:Shows how to create a list by applying a new list format to a collection of paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Paragraph 1");
        builder.writeln("Paragraph 2");
        builder.write("Paragraph 3");

        NodeCollection paras = doc.getChildNodes(NodeType.PARAGRAPH, true);

        Assert.AreEqual(0, paras.Count(n => (ms.as(n, Paragraph.class)).ListFormat.IsListItem));

        List list = doc.getLists().add(ListTemplate.NUMBER_UPPERCASE_LETTER_DOT);

        for (Paragraph paragraph : paras.<Paragraph>OfType() !!Autoporter error: Undefined expression type )
        {
            paragraph.getListFormat().setList(list);
            paragraph.getListFormat().setListLevelNumber(1);
        }

        Assert.AreEqual(3, paras.Count(n => (ms.as(n, Paragraph.class)).ListFormat.IsListItem));
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        paras = doc.getChildNodes(NodeType.PARAGRAPH, true);

        Assert.AreEqual(3, paras.Count(n => (ms.as(n, Paragraph.class)).ListFormat.IsListItem));
        Assert.AreEqual(3, paras.Count(n => (ms.as(n, Paragraph.class)).ListFormat.ListLevelNumber == 1));
    }

    //ExStart
    //ExFor:ListTemplate
    //ExSummary:Shows how to create a document that demonstrates all outline headings list templates.
    @Test //ExSkip
    public void outlineHeadingTemplates() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        List list = doc.getLists().add(ListTemplate.OUTLINE_HEADINGS_ARTICLE_SECTION);
        addOutlineHeadingParagraphs(builder, list, "Aspose.Words Outline - \"Article Section\"");

        list = doc.getLists().add(ListTemplate.OUTLINE_HEADINGS_LEGAL);
        addOutlineHeadingParagraphs(builder, list, "Aspose.Words Outline - \"Legal\"");

        builder.insertBreak(BreakType.PAGE_BREAK);

        list = doc.getLists().add(ListTemplate.OUTLINE_HEADINGS_NUMBERS);
        addOutlineHeadingParagraphs(builder, list, "Aspose.Words Outline - \"Numbers\"");

        list = doc.getLists().add(ListTemplate.OUTLINE_HEADINGS_CHAPTER);
        addOutlineHeadingParagraphs(builder, list, "Aspose.Words Outline - \"Chapters\"");

        doc.save(getArtifactsDir() + "Lists.OutlineHeadingTemplates.docx");
        testOutlineHeadingTemplates(new Document(getArtifactsDir() + "Lists.OutlineHeadingTemplates.docx")); //ExSkip
    }

    private static void addOutlineHeadingParagraphs(DocumentBuilder builder, List list, String title)
    {
        builder.getParagraphFormat().clearFormatting();
        builder.writeln(title);

        for (int i = 0; i < 9; i++)
        {
            builder.getListFormat().setList(list);
            builder.getListFormat().setListLevelNumber(i);

            String styleName = "Heading " + (i + 1);
            builder.getParagraphFormat().setStyleName(styleName);
            builder.writeln(styleName);
        }

        builder.getListFormat().removeNumbers();
    }
    //ExEnd

    private void testOutlineHeadingTemplates(Document doc)
    {
        List list = doc.getLists().get(0); // Article section list template

        TestUtil.verifyListLevel("Article \u0000.", 0.0d, NumberStyle.UPPERCASE_ROMAN, list.getListLevels().get(0));
        TestUtil.verifyListLevel("Section \u0000.\u0001", 0.0d, NumberStyle.LEADING_ZERO, list.getListLevels().get(1));
        TestUtil.verifyListLevel("(\u0002)", 14.4d, NumberStyle.LOWERCASE_LETTER, list.getListLevels().get(2));
        TestUtil.verifyListLevel("(\u0003)", 36.0d, NumberStyle.LOWERCASE_ROMAN, list.getListLevels().get(3));
        TestUtil.verifyListLevel("\u0004)", 28.8d, NumberStyle.ARABIC, list.getListLevels().get(4));
        TestUtil.verifyListLevel("\u0005)", 36.0d, NumberStyle.LOWERCASE_LETTER, list.getListLevels().get(5));
        TestUtil.verifyListLevel("\u0006)", 50.4d, NumberStyle.LOWERCASE_ROMAN, list.getListLevels().get(6));
        TestUtil.verifyListLevel("\u0007.", 50.4d, NumberStyle.LOWERCASE_LETTER, list.getListLevels().get(7));
        TestUtil.verifyListLevel("\b.", 72.0d, NumberStyle.LOWERCASE_ROMAN, list.getListLevels().get(8));

        list = doc.getLists().get(1); // Legal list template

        TestUtil.verifyListLevel("\u0000", 0.0d, NumberStyle.ARABIC, list.getListLevels().get(0));
        TestUtil.verifyListLevel("\u0000.\u0001", 0.0d, NumberStyle.ARABIC, list.getListLevels().get(1));
        TestUtil.verifyListLevel("\u0000.\u0001.\u0002", 0.0d, NumberStyle.ARABIC, list.getListLevels().get(2));
        TestUtil.verifyListLevel("\u0000.\u0001.\u0002.\u0003", 0.0d, NumberStyle.ARABIC, list.getListLevels().get(3));
        TestUtil.verifyListLevel("\u0000.\u0001.\u0002.\u0003.\u0004", 0.0d, NumberStyle.ARABIC, list.getListLevels().get(4));
        TestUtil.verifyListLevel("\u0000.\u0001.\u0002.\u0003.\u0004.\u0005", 0.0d, NumberStyle.ARABIC, list.getListLevels().get(5));
        TestUtil.verifyListLevel("\u0000.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006", 0.0d, NumberStyle.ARABIC, list.getListLevels().get(6));
        TestUtil.verifyListLevel("\u0000.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006.\u0007", 0.0d, NumberStyle.ARABIC, list.getListLevels().get(7));
        TestUtil.verifyListLevel("\u0000.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006.\u0007.\b", 0.0d, NumberStyle.ARABIC, list.getListLevels().get(8));

        list = doc.getLists().get(2); // Numbered list template

        TestUtil.verifyListLevel("\u0000.", 0.0d, NumberStyle.UPPERCASE_ROMAN, list.getListLevels().get(0));
        TestUtil.verifyListLevel("\u0001.", 36.0d, NumberStyle.UPPERCASE_LETTER, list.getListLevels().get(1));
        TestUtil.verifyListLevel("\u0002.", 72.0d, NumberStyle.ARABIC, list.getListLevels().get(2));
        TestUtil.verifyListLevel("\u0003)", 108.0d, NumberStyle.LOWERCASE_LETTER, list.getListLevels().get(3));
        TestUtil.verifyListLevel("(\u0004)", 144.0d, NumberStyle.ARABIC, list.getListLevels().get(4));
        TestUtil.verifyListLevel("(\u0005)", 180.0d, NumberStyle.LOWERCASE_LETTER, list.getListLevels().get(5));
        TestUtil.verifyListLevel("(\u0006)", 216.0d, NumberStyle.LOWERCASE_ROMAN, list.getListLevels().get(6));
        TestUtil.verifyListLevel("(\u0007)", 252.0d, NumberStyle.LOWERCASE_LETTER, list.getListLevels().get(7));
        TestUtil.verifyListLevel("(\b)", 288.0d, NumberStyle.LOWERCASE_ROMAN, list.getListLevels().get(8));

        list = doc.getLists().get(3); // Chapter list template

        TestUtil.verifyListLevel("Chapter \u0000", 0.0d, NumberStyle.ARABIC, list.getListLevels().get(0));
        TestUtil.verifyListLevel("", 0.0d, NumberStyle.NONE, list.getListLevels().get(1));
        TestUtil.verifyListLevel("", 0.0d, NumberStyle.NONE, list.getListLevels().get(2));
        TestUtil.verifyListLevel("", 0.0d, NumberStyle.NONE, list.getListLevels().get(3));
        TestUtil.verifyListLevel("", 0.0d, NumberStyle.NONE, list.getListLevels().get(4));
        TestUtil.verifyListLevel("", 0.0d, NumberStyle.NONE, list.getListLevels().get(5));
        TestUtil.verifyListLevel("", 0.0d, NumberStyle.NONE, list.getListLevels().get(6));
        TestUtil.verifyListLevel("", 0.0d, NumberStyle.NONE, list.getListLevels().get(7));
        TestUtil.verifyListLevel("", 0.0d, NumberStyle.NONE, list.getListLevels().get(8));
    }

    //ExStart
    //ExFor:ListCollection
    //ExFor:ListCollection.AddCopy(List)
    //ExFor:ListCollection.GetEnumerator
    //ExSummary:Shows how to enumerate through all lists defined in one document and creates a sample of those lists in another document.
    @Test //ExSkip
    public void printOutAllLists() throws Exception
    {
        // Open a document that contains lists
        Document srcDoc = new Document(getMyDir() + "Rendering.docx");

        // This will be the sample document we product
        Document dstDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        for (List srcList : srcDoc.getLists())
        {
            // This copies the list formatting from the source into the destination document
            List dstList = dstDoc.getLists().addCopy(srcList);
            addListSample(builder, dstList);
        }

        dstDoc.save(getArtifactsDir() + "Lists.PrintOutAllLists.docx");
        testPrintOutAllLists(srcDoc, new Document(getArtifactsDir() + "Lists.PrintOutAllLists.docx")); //ExSkip
    }

    private static void addListSample(DocumentBuilder builder, List list)
    {
        builder.writeln("Sample formatting of list with ListId:" + list.getListId());
        builder.getListFormat().setList(list);
        for (int i = 0; i < list.getListLevels().getCount(); i++)
        {
            builder.getListFormat().setListLevelNumber(i);
            builder.writeln("Level " + i);
        }

        builder.getListFormat().removeNumbers();
        builder.writeln();
    }
    //ExEnd		

    private void testPrintOutAllLists(Document listSourceDoc, Document outDoc)
    {
        for (List list : outDoc.getLists())
            for (int i = 0; i < list.getListLevels().getCount(); i++)
            {
                ListLevel expectedListLevel = listSourceDoc.getLists().First(l => l.ListId == list.ListId).ListLevels[i];
                Assert.assertEquals(expectedListLevel.getNumberFormat(), list.getListLevels().get(i).getNumberFormat());
                Assert.assertEquals(expectedListLevel.getNumberPosition(), list.getListLevels().get(i).getNumberPosition());
                Assert.assertEquals(expectedListLevel.getNumberStyle(), list.getListLevels().get(i).getNumberStyle());
            }
    }

    @Test
    public void listDocument() throws Exception
    {
        //ExStart
        //ExFor:ListCollection.Document
        //ExFor:ListCollection.Count
        //ExFor:ListCollection.Item(Int32)
        //ExFor:ListCollection.GetListByListId
        //ExFor:List.Document
        //ExFor:List.ListId
        //ExSummary:Shows how to verify owner document properties of lists.
        Document doc = new Document();

        ListCollection lists = doc.getLists();

        Assert.assertEquals(doc, lists.getDocument());

        List list = lists.add(ListTemplate.BULLET_DEFAULT);

        Assert.assertEquals(doc, list.getDocument());

        System.out.println("Current list count: " + lists.getCount());
        System.out.println("Is the first document list: " + (lists.get(0).equals(list)));
        System.out.println("ListId: " + list.getListId());
        System.out.println("List is the same by ListId: " + (lists.getListByListId(1).equals(list)));
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        lists = doc.getLists();
        
        Assert.assertEquals(doc, lists.getDocument());
        Assert.assertEquals(1, lists.getCount());
        Assert.assertEquals(1, lists.get(0).getListId());
        Assert.assertEquals(lists.get(0), lists.getListByListId(1));
    }
    
    @Test
    public void createListRestartAfterHigher() throws Exception
    {
        //ExStart
        //ExFor:ListLevel.NumberStyle
        //ExFor:ListLevel.NumberFormat
        //ExFor:ListLevel.IsLegal
        //ExFor:ListLevel.RestartAfterLevel
        //ExFor:ListLevel.LinkedStyle
        //ExFor:ListLevelCollection.GetEnumerator
        //ExSummary:Shows how to create a list with some advanced formatting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        List list = doc.getLists().add(ListTemplate.NUMBER_DEFAULT);

        // Level 1 labels will be "Appendix A", continuous and linked to the Heading 1 paragraph style
        list.getListLevels().get(0).setNumberFormat("Appendix \u0000");
        list.getListLevels().get(0).setNumberStyle(NumberStyle.UPPERCASE_LETTER);
        list.getListLevels().get(0).setLinkedStyle(doc.getStyles().get("Heading 1"));

        // Level 2 labels will be "Section (1.01)" and restarting after Level 2 item appears
        list.getListLevels().get(1).setNumberFormat("Section (\u0000.\u0001)");
        list.getListLevels().get(1).setNumberStyle(NumberStyle.LEADING_ZERO);
        // Notice the higher level uses UppercaseLetter numbering, but we want arabic number
        // of the higher levels to appear in this level, therefore set this property
        list.getListLevels().get(1).isLegal(true);
        list.getListLevels().get(1).setRestartAfterLevel(0);

        // Level 3 labels will be "-I-" and restarting after Level 2 item appears
        list.getListLevels().get(2).setNumberFormat("-\u0002-");
        list.getListLevels().get(2).setNumberStyle(NumberStyle.UPPERCASE_ROMAN);
        list.getListLevels().get(2).setRestartAfterLevel(1);

        // Make labels of all list levels bold
        for (ListLevel level : list.getListLevels())
            level.getFont().setBold(true);

        // Apply list formatting to the current paragraph
        builder.getListFormat().setList(list);

        // Exercise the 3 levels we created two times
        for (int n = 0; n < 2; n++)
        {
            for (int i = 0; i < 3; i++)
            {
                builder.getListFormat().setListLevelNumber(i);
                builder.writeln("Level " + i);
            }
        }

        builder.getListFormat().removeNumbers();

        doc.save(getArtifactsDir() + "Lists.CreateListRestartAfterHigher.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Lists.CreateListRestartAfterHigher.docx");

        ListLevel listLevel = doc.getLists().get(0).getListLevels().get(0);

        TestUtil.verifyListLevel("Appendix \u0000", 18.0d, NumberStyle.UPPERCASE_LETTER, listLevel);
        Assert.assertFalse(listLevel.isLegal());
        Assert.assertEquals(-1, listLevel.getRestartAfterLevel());
        Assert.assertEquals("Heading 1", listLevel.getLinkedStyle().getName());

        listLevel = doc.getLists().get(0).getListLevels().get(1);

        TestUtil.verifyListLevel("Section (\u0000.\u0001)", 54.0d, NumberStyle.LEADING_ZERO, listLevel);
        Assert.assertTrue(listLevel.isLegal());
        Assert.assertEquals(0, listLevel.getRestartAfterLevel());
        Assert.assertNull(listLevel.getLinkedStyle());
    }

    @Test
    public void getListLabels() throws Exception
    {
        //ExStart
        //ExFor:Document.UpdateListLabels()
        //ExFor:Node.ToString(SaveFormat)
        //ExFor:ListLabel
        //ExFor:Paragraph.ListLabel
        //ExFor:ListLabel.LabelValue
        //ExFor:ListLabel.LabelString
        //ExSummary:Shows how to extract the label of each paragraph in a list as a value or a String.
        Document doc = new Document(getMyDir() + "Rendering.docx");
        doc.updateListLabels();

        NodeCollection paras = doc.getChildNodes(NodeType.PARAGRAPH, true);

        // Find if we have the paragraph list. In our document our list uses plain arabic numbers,
        // which start at three and ends at six
        for (Paragraph paragraph : paras.<Paragraph>OfType().Where(p => p.ListFormat.IsListItem) !!Autoporter error: Undefined expression type )
        {
            System.out.println("List item paragraph #{paras.IndexOf(paragraph)}");

            // This is the text we get when actually getting when we output this node to text format
            // The list labels are not included in this text output. Trim any paragraph formatting characters
            String paragraphText = msString.trim(paragraph.toString(SaveFormat.TEXT));
            System.out.println("\tExported Text: {paragraphText}");

            ListLabel label = paragraph.getListLabel();
            // This gets the position of the paragraph in current level of the list. If we have a list with multiple level then this
            // will tell us what position it is on that particular level
            System.out.println("\tNumerical Id: {label.LabelValue}");

            // Combine them together to include the list label with the text in the output
            System.out.println("\tList label combined with text: {label.LabelString} {paragraphText}");
        }
        //ExEnd

        Assert.AreEqual(10, paras.<Paragraph>OfType().Count(p => p.ListFormat.IsListItem));
    }

    @Test
    public void createPictureBullet() throws Exception
    {
        //ExStart
        //ExFor:ListLevel.CreatePictureBullet
        //ExFor:ListLevel.DeletePictureBullet
        //ExSummary:Shows how to creating and deleting picture bullet with custom image.
        Document doc = new Document();

        // Create a list with template
        List list = doc.getLists().add(ListTemplate.BULLET_CIRCLE);

        // Create picture bullet for the current list level
        list.getListLevels().get(0).createPictureBullet();

        // Set your own picture bullet image through the ImageData
        list.getListLevels().get(0).getImageData().setImage(getImageDir() + "Logo icon.ico");

        Assert.assertTrue(list.getListLevels().get(0).getImageData().hasImage());

        // Create a list, configure its bullets to use our image and add two list items
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getListFormat().setList(list);
        builder.writeln("Hello world!");
        builder.write("Hello again!");

        doc.save(getArtifactsDir() + "Lists.CreatePictureBullet.docx");

        // Delete picture bullet
        list.getListLevels().get(0).deletePictureBullet();

        Assert.assertNull(list.getListLevels().get(0).getImageData());
        //ExEnd

        doc = new Document(getArtifactsDir() + "Lists.CreatePictureBullet.docx");

        Assert.assertTrue(doc.getLists().get(0).getListLevels().get(0).getImageData().hasImage());
    }
}
