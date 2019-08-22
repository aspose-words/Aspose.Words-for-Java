// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Document;
import com.aspose.words.ListTemplate;
import com.aspose.words.List;
import com.aspose.words.StyleIdentifier;
import org.testng.Assert;
import com.aspose.words.ListLevel;
import java.awt.Color;
import com.aspose.words.NumberStyle;
import com.aspose.words.ListLevelAlignment;
import com.aspose.words.ListTrailingCharacter;
import com.aspose.words.Style;
import com.aspose.words.StyleType;
import com.aspose.ms.System.msConsole;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.ms.System.msString;
import com.aspose.words.Body;
import com.aspose.words.BreakType;
import com.aspose.words.ListCollection;
import com.aspose.words.SaveFormat;
import com.aspose.words.ListLabel;


@Test
public class ExLists extends ApiExampleBase
{
    private /*final*/ String _image = getImageDir() + "Test_636_852.gif";

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
        //ExSummary:Shows how to apply default bulleted or numbered list formatting to paragraphs when using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder();

        builder.writeln("Aspose.Words allows:");
        builder.writeln();

        // Start a numbered list with default formatting.
        builder.getListFormat().applyNumberDefault();
        builder.writeln("Opening documents from different formats:");

        // Go to second list level, add more text.
        builder.getListFormat().listIndent();
        builder.writeln("DOC");
        builder.writeln("PDF");
        builder.writeln("HTML");

        // Outdent to the first list level.
        builder.getListFormat().listOutdent();
        builder.writeln("Processing documents");
        builder.writeln("Saving documents in different formats:");

        // Indent the list level again.
        builder.getListFormat().listIndent();
        builder.writeln("DOC");
        builder.writeln("PDF");
        builder.writeln("HTML");
        builder.writeln("MHTML");
        builder.writeln("Plain text");

        // Outdent the list level again.
        builder.getListFormat().listOutdent();
        builder.writeln("Doing many other things!");

        // End the numbered list.
        builder.getListFormat().removeNumbers();
        builder.writeln();

        builder.writeln("Aspose.Words main advantages are:");
        builder.writeln();

        // Start a bulleted list with default formatting.
        builder.getListFormat().applyBulletDefault();
        builder.writeln("Great performance");
        builder.writeln("High reliability");
        builder.writeln("Quality code and working");
        builder.writeln("Wide variety of features");
        builder.writeln("Easy to understand API");

        // End the bulleted list.
        builder.getListFormat().removeNumbers();

        builder.getDocument().save(getArtifactsDir() + "Lists.ApplyDefaultBulletsAndNumbers.doc");
        //ExEnd
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
        // apply it to the current paragraph in the document builder.
        builder.getListFormat().setList(doc.getLists().add(ListTemplate.NUMBER_ARABIC_DOT));

        // There are 9 levels in this list, lets try them all.
        for (int i = 0; i < 9; i++)
        {
            builder.getListFormat().setListLevelNumber(i);
            builder.writeln("Level " + i);
        }

        // Create a bulleted list based on one of the Microsoft Word list templates
        // and apply it to the current paragraph in the document builder.
        builder.getListFormat().setList(doc.getLists().add(ListTemplate.BULLET_DIAMONDS));

        // There are 9 levels in this list, lets try them all.
        for (int i = 0; i < 9; i++)
        {
            builder.getListFormat().setListLevelNumber(i);
            builder.writeln("Level " + i);
        }

        // This is a way to stop list formatting. 
        builder.getListFormat().setList(null);

        builder.getDocument().save(getArtifactsDir() + "Lists.SpecifyListLevel.doc");
        //ExEnd
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

        // Create an outline list for the headings.
        List outlineList = doc.getLists().add(ListTemplate.OUTLINE_NUMBERS);
        builder.getListFormat().setList(outlineList);
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.writeln("This is my Chapter 1");

        // Create a numbered list.
        List numberedList = doc.getLists().add(ListTemplate.NUMBER_DEFAULT);
        builder.getListFormat().setList(numberedList);
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.NORMAL);
        builder.writeln("Numbered list item 1.");

        // Every paragraph that comprises a list will have this flag
        Assert.assertTrue(builder.getCurrentParagraph().isListItem());
        Assert.assertTrue(builder.getParagraphFormat().isListItem());

        // Create a bulleted list.
        List bulletedList = doc.getLists().add(ListTemplate.BULLET_DEFAULT);
        builder.getListFormat().setList(bulletedList);
        builder.getParagraphFormat().setLeftIndent(72.0);
        builder.writeln("Bulleted list item 1.");
        builder.writeln("Bulleted list item 2.");
        builder.getParagraphFormat().clearFormatting();

        // Revert to the numbered list.
        builder.getListFormat().setList(numberedList);
        builder.writeln("Numbered list item 2.");
        builder.writeln("Numbered list item 3.");

        // Revert to the outline list.
        builder.getListFormat().setList(outlineList);
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.writeln("This is my Chapter 2");

        builder.getParagraphFormat().clearFormatting();

        builder.getDocument().save(getArtifactsDir() + "Lists.NestedLists.docx");
        //ExEnd
    }

    @Test
    public void createCustomList() throws Exception
    {
        //ExStart
        //ExFor:List
        //ExFor:List.ListLevels
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

        // Create a list based on one of the Microsoft Word list templates.
        List list = doc.getLists().add(ListTemplate.NUMBER_DEFAULT);

        // Completely customize one list level.
        ListLevel level1 = list.getListLevels().get(0);
        level1.getFont().setColor(Color.RED);
        level1.getFont().setSize(24.0);
        level1.setNumberStyle(NumberStyle.ORDINAL_TEXT);
        level1.setStartAt(21);
        level1.setNumberFormat("\u0000");

        level1.setNumberPosition(-36);
        level1.setTextPosition(144.0);
        level1.setTabPosition(144.0);

        // Completely customize yet another list level.
        ListLevel level2 = list.getListLevels().get(1);
        level2.setAlignment(ListLevelAlignment.RIGHT);
        level2.setNumberStyle(NumberStyle.BULLET);
        level2.getFont().setName("Wingdings");
        level2.getFont().setColor(Color.BLUE);
        level2.getFont().setSize(24.0);
        level2.setNumberFormat("\uf0af"); // A bullet that looks like some sort of a star.
        level2.setTrailingCharacter(ListTrailingCharacter.SPACE);
        level2.setNumberPosition(144.0);

        // Now add some text that uses the list that we created.			
        // It does not matter when to customize the list - before or after adding the paragraphs.
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

        builder.getDocument().save(getArtifactsDir() + "Lists.CreateCustomList.doc");
        //ExEnd
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
        //ExFor:ListFormat.List
        //ExSummary:Shows how to restart numbering in a list by copying a list.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list based on a template.
        List list1 = doc.getLists().add(ListTemplate.NUMBER_ARABIC_PARENTHESIS);
        // Modify the formatting of the list.
        list1.getListLevels().get(0).getFont().setColor(Color.RED);
        list1.getListLevels().get(0).setAlignment(ListLevelAlignment.RIGHT);

        builder.writeln("List 1 starts below:");
        // Use the first list in the document for a while.
        builder.getListFormat().setList(list1);
        builder.writeln("Item 1");
        builder.writeln("Item 2");
        builder.getListFormat().removeNumbers();

        // Now I want to reuse the first list, but need to restart numbering.
        // This should be done by creating a copy of the original list formatting.
        List list2 = doc.getLists().addCopy(list1);

        // We can modify the new list in any way. Including setting new start number.
        list2.getListLevels().get(0).setStartAt(10);

        // Use the second list in the document.
        builder.writeln("List 2 starts below:");
        builder.getListFormat().setList(list2);
        builder.writeln("Item 1");
        builder.writeln("Item 2");
        builder.getListFormat().removeNumbers();

        builder.getDocument().save(getArtifactsDir() + "Lists.RestartNumberingUsingListCopy.doc");
        //ExEnd
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

        // Create a new list style. 
        // List formatting associated with this list style is default numbered.
        Style listStyle = doc.getStyles().add(StyleType.LIST, "MyListStyle");

        // This list defines the formatting of the list style.
        // Note this list can not be used directly to apply formatting to paragraphs (see below).
        List list1 = listStyle.getList();

        // Check some basic rules about the list that defines a list style.
        msConsole.writeLine("IsListStyleDefinition: " + list1.isListStyleDefinition());
        msConsole.writeLine("IsListStyleReference: " + list1.isListStyleReference());
        msConsole.writeLine("IsMultiLevel: " + list1.isMultiLevel());
        msConsole.writeLine("List style has been set: " + (listStyle == list1.getStyle()));

        // Modify formatting of the list style to our liking.
        for (int i = 0; i < list1.getListLevels().getCount(); i++)
        {
            ListLevel level = list1.getListLevels().get(i);
            level.getFont().setName("Verdana");
            level.getFont().setColor(Color.BLUE);
            level.getFont().setBold(true);
        }

        // Add some text to our document and use the list style.
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Using list style first time:");

        // This creates a list based on the list style.
        List list2 = doc.getLists().add(listStyle);

        // Check some basic rules about the list that references a list style.
        msConsole.writeLine("IsListStyleDefinition: " + list2.isListStyleDefinition());
        msConsole.writeLine("IsListStyleReference: " + list2.isListStyleReference());
        msConsole.writeLine("List Style has been set: " + (listStyle == list2.getStyle()));

        // Apply the list that references the list style.
        builder.getListFormat().setList(list2);
        builder.writeln("Item 1");
        builder.writeln("Item 2");
        builder.getListFormat().removeNumbers();

        builder.writeln("Using list style second time:");

        // Create and apply another list based on the list style.
        List list3 = doc.getLists().add(listStyle);
        builder.getListFormat().setList(list3);
        builder.writeln("Item 1");
        builder.writeln("Item 2");
        builder.getListFormat().removeNumbers();

        builder.getDocument().save(getArtifactsDir() + "Lists.CreateAndUseListStyle.doc");
        //ExEnd

        // Verify properties of list 1
        Assert.assertTrue(list1.isListStyleDefinition());
        Assert.assertFalse(list1.isListStyleReference());
        Assert.assertTrue(list1.isMultiLevel());
        msAssert.areEqual(listStyle, list1.getStyle());

        // Verify properties of list 2
        Assert.assertFalse(list2.isListStyleDefinition());
        Assert.assertTrue(list2.isListStyleReference());
        msAssert.areEqual(listStyle, list2.getStyle());
    }

    @Test
    public void detectBulletedParagraphs() throws Exception
    {
        Document doc = new Document();

        //ExStart
        //ExFor:Paragraph.ListFormat
        //ExFor:ListFormat.IsListItem
        //ExFor:CompositeNode.GetText
        //ExFor:List.ListId
        //ExSummary:Finds and outputs all paragraphs in a document that are bulleted or numbered.
        NodeCollection paras = doc.getChildNodes(NodeType.PARAGRAPH, true);
        for (Paragraph para : paras.<Paragraph>OfType() !!Autoporter error: Undefined expression type )
        {
            if (para.getListFormat().isListItem())
            {
                msConsole.writeLine(msString.format("*** A paragraph belongs to list {0}",
                    para.getListFormat().getList().getListId()));
                msConsole.writeLine(para.getText());
            }
        }

        //ExEnd
    }

    @Test
    public void removeBulletsFromParagraphs() throws Exception
    {
        Document doc = new Document();

        //ExStart
        //ExFor:Paragraph.ListFormat
        //ExFor:ListFormat.RemoveNumbers
        //ExSummary:Removes bullets and numbering from all paragraphs in the main text of a section.
        Body body = doc.getFirstSection().getBody();

        for (Paragraph paragraph : body.getParagraphs().<Paragraph>OfType() !!Autoporter error: Undefined expression type )
            paragraph.getListFormat().removeNumbers();

        //ExEnd
    }

    @Test
    public void applyExistingListToParagraphs() throws Exception
    {
        Document doc = new Document();
        doc.getLists().add(ListTemplate.NUMBER_DEFAULT);

        //ExStart
        //ExFor:Paragraph.ListFormat
        //ExFor:ListFormat.List
        //ExFor:ListFormat.ListLevelNumber
        //ExFor:ListCollection.Item(Int32)
        //ExSummary:Applies list formatting of an existing list to a collection of paragraphs.
        Body body = doc.getFirstSection().getBody();
        List list = doc.getLists().get(0);
        for (Paragraph paragraph : body.getParagraphs().<Paragraph>OfType() !!Autoporter error: Undefined expression type )
        {
            paragraph.getListFormat().setList(list);
            paragraph.getListFormat().setListLevelNumber(2);
        }

        //ExEnd
    }

    @Test
    public void applyNewListToParagraphs() throws Exception
    {
        Document doc = new Document();

        //ExStart
        //ExFor:Paragraph.ListFormat
        //ExFor:ListFormat.ListLevelNumber
        //ExFor:ListCollection.Add(ListTemplate)
        //ExSummary:Creates new list formatting and applies it to a collection of paragraphs.
        List list = doc.getLists().add(ListTemplate.NUMBER_UPPERCASE_LETTER_DOT);

        Body body = doc.getFirstSection().getBody();
        for (Paragraph paragraph : body.getParagraphs().<Paragraph>OfType() !!Autoporter error: Undefined expression type )
        {
            paragraph.getListFormat().setList(list);
            paragraph.getListFormat().setListLevelNumber(1);
        }

        //ExEnd
    }

    /// <summary>
    /// This calls the below method to resolve skipping of [Test] in VB.NET.
    /// </summary>
    @Test
    public void outlineHeadingTemplatesCaller() throws Exception
    {
        outlineHeadingTemplates();
    }

    //ExStart
    //ExFor:ListTemplate
    //ExSummary:Creates a sample document that exercises all outline headings list templates.
    @Test (enabled = false)
    public void outlineHeadingTemplates() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        List list = doc.getLists().add(ListTemplate.OUTLINE_HEADINGS_ARTICLE_SECTION);
        addOutlineHeadingParagraphs(builder, list, "Aspose.Words Outline 1");

        list = doc.getLists().add(ListTemplate.OUTLINE_HEADINGS_LEGAL);
        addOutlineHeadingParagraphs(builder, list, "Aspose.Words Outline 2");

        builder.insertBreak(BreakType.PAGE_BREAK);

        list = doc.getLists().add(ListTemplate.OUTLINE_HEADINGS_NUMBERS);
        addOutlineHeadingParagraphs(builder, list, "Aspose.Words Outline 3");

        list = doc.getLists().add(ListTemplate.OUTLINE_HEADINGS_CHAPTER);
        addOutlineHeadingParagraphs(builder, list, "Aspose.Words Outline 4");

        builder.getDocument().save(getArtifactsDir() + "Lists.OutlineHeadingTemplates.doc");
    }

    private static void addOutlineHeadingParagraphs(DocumentBuilder builder, List list, String title)
    {
        builder.getParagraphFormat().clearFormatting();
        builder.writeln(title);

        for (int i = 0; i < 9; i++)
        {
            builder.getListFormat().setList(list);
            builder.getListFormat().setListLevelNumber(i);

            String styleName = "Heading " + Integer.toString((i + 1));
            builder.getParagraphFormat().setStyleName(styleName);
            builder.writeln(styleName);
        }

        builder.getListFormat().removeNumbers();
    }
    //ExEnd

    @Test
    public void printOutAllLists() throws Exception
    {
        //ExStart
        //ExFor:ListCollection
        //ExFor:ListCollection.AddCopy(List)
        //ExFor:ListCollection.GetEnumerator
        //ExSummary:Enumerates through all lists defined in one document and creates a sample of those lists in another document.
        // You can use any of your documents to try this little program out.
        Document srcDoc = new Document(getMyDir() + "Lists.PrintOutAllLists.doc");

        // This will be the sample document we product.
        Document dstDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        for (List srcList : srcDoc.getLists())
        {
            // This copies the list formatting from the source into the destination document.
            List dstList = dstDoc.getLists().addCopy(srcList);
            addListSample(builder, dstList);
        }

        dstDoc.save(getArtifactsDir() + "Lists.PrintOutAllLists.doc");
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
        //ExSummary:Illustrates the owner document properties of lists.
        Document doc = new Document();

        ListCollection lists = doc.getLists();
        // All of these should be equal.
        msConsole.writeLine("ListCollection document is doc: " + (doc == lists.getDocument()));
        msConsole.writeLine("Starting list count: " + lists.getCount());

        List list = lists.add(ListTemplate.BULLET_DEFAULT);
        msConsole.writeLine("List document is doc: " + (list.getDocument() == doc));
        msConsole.writeLine("List count after adding list: " + lists.getCount());
        msConsole.writeLine("Is the first document list: " + (lists.get(0).equals(list)));
        msConsole.writeLine("ListId: " + list.getListId());
        msConsole.writeLine("List is the same by ListId: " + (lists.getListByListId(1).equals(list)));
        //ExEnd

        // Verify these properties
        msAssert.areEqual(doc, lists.getDocument());
        msAssert.areEqual(doc, list.getDocument());
        msAssert.areEqual(1, lists.getCount());
        msAssert.areEqual(list, lists.get(0));
        msAssert.areEqual(1, list.getListId());
        msAssert.areEqual(list, lists.getListByListId(1));
    }

    @Test
    public void listFormatListLevel() throws Exception
    {
        //ExStart
        //ExFor:ListFormat.ListLevel
        //ExSummary:Shows how to modify list formatting of the current list level.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create and apply list formatting to the current paragraph.
        builder.getListFormat().setList(doc.getLists().add(ListTemplate.NUMBER_DEFAULT));

        // Modify formatting of the current (first) list level.
        builder.getListFormat().getListLevel().getFont().setBold(true);

        builder.writeln("Item 1");
        builder.writeln("Item 2");
        builder.getListFormat().removeNumbers();
        //ExEnd
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

        // Level 1 labels will be "Appendix A", continuous and linked to the Heading 1 paragraph style.
        list.getListLevels().get(0).setNumberFormat("Appendix \u0000");
        list.getListLevels().get(0).setNumberStyle(NumberStyle.UPPERCASE_LETTER);
        list.getListLevels().get(0).setLinkedStyle(doc.getStyles().get("Heading 1"));

        // Level 2 labels will be "Section (1.01)" and restarting after Level 2 item appears.
        list.getListLevels().get(1).setNumberFormat("Section (\u0000.\u0001)");
        list.getListLevels().get(1).setNumberStyle(NumberStyle.LEADING_ZERO);
        // Notice the higher level uses UppercaseLetter numbering, but we want arabic number
        // of the higher levels to appear in this level, therefore set this property.
        list.getListLevels().get(1).isLegal(true);
        list.getListLevels().get(1).setRestartAfterLevel(0);

        // Level 3 labels will be "-I-" and restarting after Level 2 item appears.
        list.getListLevels().get(2).setNumberFormat("-\u0002-");
        list.getListLevels().get(2).setNumberStyle(NumberStyle.UPPERCASE_ROMAN);
        list.getListLevels().get(2).setRestartAfterLevel(1);

        // Make labels of all list levels bold.
        for (ListLevel level : list.getListLevels())
            level.getFont().setBold(true);

        // Apply list formatting to the current paragraph.
        builder.getListFormat().setList(list);

        // Exercise the 3 levels we created two times.
        for (int n = 0; n < 2; n++)
        {
            for (int i = 0; i < 3; i++)
            {
                builder.getListFormat().setListLevelNumber(i);
                builder.writeln("Level " + i);
            }
        }

        builder.getListFormat().removeNumbers();

        builder.getDocument().save(getArtifactsDir() + "Lists.CreateListRestartAfterHigher.doc");
        //ExEnd
    }

    @Test
    public void paragraphStyleBulleted() throws Exception
    {
        //ExStart
        //ExFor:StyleCollection
        //ExFor:DocumentBase.Styles
        //ExFor:Style
        //ExFor:Font
        //ExFor:Style.Font
        //ExFor:Style.ParagraphFormat
        //ExFor:Style.ListFormat
        //ExFor:ParagraphFormat.Style
        //ExSummary:Shows how to create and use a paragraph style with list formatting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a paragraph style and specify some formatting for it.
        Style style = doc.getStyles().add(StyleType.PARAGRAPH, "MyStyle1");
        style.getFont().setSize(24.0);
        style.getFont().setName("Verdana");
        style.getParagraphFormat().setSpaceAfter(12.0);

        // Create a list and make sure the paragraphs that use this style will use this list.
        style.getListFormat().setList(doc.getLists().add(ListTemplate.BULLET_DEFAULT));
        style.getListFormat().setListLevelNumber(0);

        // Apply the paragraph style to the current paragraph in the document and add some text.
        builder.getParagraphFormat().setStyle(style);
        builder.writeln("Hello World: MyStyle1, bulleted.");

        // Change to a paragraph style that has no list formatting.
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Normal"));
        builder.writeln("Hello World: Normal.");

        builder.getDocument().save(getArtifactsDir() + "Lists.ParagraphStyleBulleted.doc");
        //ExEnd
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
        Document doc = new Document(getMyDir() + "Lists.PrintOutAllLists.doc");
        doc.updateListLabels();
        int listParaCount = 1;

        for (Paragraph paragraph : doc.getChildNodes(NodeType.PARAGRAPH, true).<Paragraph>OfType() !!Autoporter error: Undefined expression type )
        {
            // Find if we have the paragraph list. In our document our list uses plain arabic numbers,
            // which start at three and ends at six.
            if (paragraph.getListFormat().isListItem())
            {
                msConsole.writeLine("Paragraph #{0}", listParaCount);

                // This is the text we get when actually getting when we output this node to text format. 
                // The list labels are not included in this text output. Trim any paragraph formatting characters.
                String paragraphText = msString.trim(paragraph.toString(SaveFormat.TEXT));
                msConsole.writeLine("Exported Text: " + paragraphText);

                ListLabel label = paragraph.getListLabel();
                // This gets the position of the paragraph in current level of the list. If we have a list with multiple level then this
                // will tell us what position it is on that particular level.
                msConsole.writeLine("Numerical Id: " + label.getLabelValue());

                // Combine them together to include the list label with the text in the output.
                msConsole.writeLine("List label combined with text: " + label.getLabelString() + " " + paragraphText);

                listParaCount++;
            }
        }

        //ExEnd
    }

    @Test
    public void createPictureBullet() throws Exception
    {
        //ExStart
        //ExFor:ListLevel.CreatePictureBullet
        //ExFor:ListLevel.DeletePictureBullet
        //ExSummary:Shows how to creating and deleting picture bullet with custom image
        Document doc = new Document();

        // Create a list with template
        List list = doc.getLists().add(ListTemplate.BULLET_CIRCLE);

        // Create picture bullet for the current list level
        list.getListLevels().get(0).createPictureBullet();

        // Set your own picture bullet image through the ImageData
        list.getListLevels().get(0).getImageData().setImage(_image);

        Assert.assertTrue(list.getListLevels().get(0).getImageData().hasImage());

        // Delete picture bullet
        list.getListLevels().get(0).deletePictureBullet();

        Assert.assertNull(list.getListLevels().get(0).getImageData());
        //ExEnd
    }
}
