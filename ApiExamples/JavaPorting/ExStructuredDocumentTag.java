// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
import java.util.ArrayList;
import com.aspose.words.StructuredDocumentTag;
import com.aspose.words.NodeType;
import org.testng.Assert;
import com.aspose.words.SdtType;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Style;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.MarkupLevel;
import com.aspose.words.NodeCollection;
import com.aspose.words.Node;
import com.aspose.ms.System.Globalization.msCultureInfo;
import com.aspose.words.SdtDateStorageFormat;
import com.aspose.words.SdtCalendarType;
import com.aspose.ms.System.DateTime;
import java.awt.Color;
import com.aspose.words.GlossaryDocument;
import com.aspose.words.BuildingBlock;
import com.aspose.words.Section;
import com.aspose.words.Body;
import com.aspose.words.SdtListItemCollection;
import com.aspose.words.SdtListItem;
import java.util.Iterator;
import com.aspose.ms.System.msConsole;
import com.aspose.ms.System.Guid;
import com.aspose.words.CustomXmlPart;
import com.aspose.ms.System.Text.Encoding;
import com.aspose.words.CustomXmlPartCollection;
import com.aspose.words.StructuredDocumentTagRangeStart;
import com.aspose.words.CustomXmlSchemaCollection;
import com.aspose.words.Run;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.Table;
import com.aspose.words.Row;
import com.aspose.words.StructuredDocumentTagRangeEnd;
import org.testng.annotations.DataProvider;
import com.aspose.words.ref.Ref;


@Test
class ExStructuredDocumentTag !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void repeatingSection() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.SdtType
        //ExSummary:Shows how to get the type of a structured document tag.
        Document doc = new Document(getMyDir() + "Structured document tags.docx");

        ArrayList<StructuredDocumentTag> sdTags = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true).<StructuredDocumentTag>OfType().ToList();

        Assert.assertEquals(SdtType.REPEATING_SECTION, sdTags.get(0).getSdtType());
        Assert.assertEquals(SdtType.REPEATING_SECTION_ITEM, sdTags.get(1).getSdtType());
        Assert.assertEquals(SdtType.RICH_TEXT, sdTags.get(2).getSdtType());
        //ExEnd
    }

    @Test
    public void applyStyle() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag
        //ExFor:StructuredDocumentTag.NodeType
        //ExFor:StructuredDocumentTag.Style
        //ExFor:StructuredDocumentTag.StyleName
        //ExFor:MarkupLevel
        //ExFor:SdtType
        //ExSummary:Shows how to work with styles for content control elements.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Below are two ways to apply a style from the document to a structured document tag.
        // 1 -  Apply a style object from the document's style collection:
        Style quoteStyle = doc.getStyles().getByStyleIdentifier(StyleIdentifier.QUOTE);
        StructuredDocumentTag sdtPlainText = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.INLINE);
        sdtPlainText.setStyle(quoteStyle);

        // 2 -  Reference a style in the document by name:
        StructuredDocumentTag sdtRichText = new StructuredDocumentTag(doc, SdtType.RICH_TEXT, MarkupLevel.INLINE);
        sdtRichText.setStyleName("Quote");

        builder.insertNode(sdtPlainText);
        builder.insertNode(sdtRichText);

        Assert.assertEquals(NodeType.STRUCTURED_DOCUMENT_TAG, sdtPlainText.getNodeType());

        NodeCollection tags = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        for (Node node : (Iterable<Node>) tags)
        {
            StructuredDocumentTag sdt = (StructuredDocumentTag)node;

            Assert.assertEquals(StyleIdentifier.QUOTE, sdt.getStyle().getStyleIdentifier());
            Assert.assertEquals("Quote", sdt.getStyleName());
        }
        //ExEnd
    }

    @Test
    public void checkBox() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.#ctor(DocumentBase, SdtType, MarkupLevel)
        //ExFor:StructuredDocumentTag.Checked
        //ExSummary:Show how to create a structured document tag in the form of a check box.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        StructuredDocumentTag sdtCheckBox = new StructuredDocumentTag(doc, SdtType.CHECKBOX, MarkupLevel.INLINE);
        sdtCheckBox.setChecked(true);

        builder.insertNode(sdtCheckBox);

        doc.save(getArtifactsDir() + "StructuredDocumentTag.CheckBox.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "StructuredDocumentTag.CheckBox.docx");

        StructuredDocumentTag[] sdts = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true).<StructuredDocumentTag>OfType().ToArray();

        Assert.assertEquals(true, sdts[0].getChecked());
        Assert.That(sdts[0].getXmlMapping().getStoreItemId(), Is.Empty);
    }

    @Test (groups = "SkipMono")
    public void date() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.CalendarType
        //ExFor:StructuredDocumentTag.DateDisplayFormat
        //ExFor:StructuredDocumentTag.DateDisplayLocale
        //ExFor:StructuredDocumentTag.DateStorageFormat
        //ExFor:StructuredDocumentTag.FullDate
        //ExSummary:Shows how to prompt the user to enter a date with a structured document tag.
        Document doc = new Document();

        // Insert a structured document tag that prompts the user to enter a date.
        // In Microsoft Word, this element is known as a "Date picker content control".
        // When we click on the arrow on the right end of this tag in Microsoft Word,
        // we will see a pop up in the form of a clickable calendar.
        // We can use that popup to select a date that the tag will display.
        StructuredDocumentTag sdtDate = new StructuredDocumentTag(doc, SdtType.DATE, MarkupLevel.INLINE);

        // Display the date, according to the Saudi Arabian Arabic locale.
        sdtDate.setDateDisplayLocale(msCultureInfo.getCultureInfo("ar-SA").getLCID());

        // Set the format with which to display the date.
        sdtDate.setDateDisplayFormat("dd MMMM, yyyy");
        sdtDate.setDateStorageFormat(SdtDateStorageFormat.DATE_TIME);

        // Display the date according to the Hijri calendar.
        sdtDate.setCalendarType(SdtCalendarType.HIJRI);

        // Before the user chooses a date in Microsoft Word, the tag will display the text "Click here to enter a date.".
        // According to the tag's calendar, set the "FullDate" property to get the tag to display a default date.
        sdtDate.setFullDateInternal(new DateTime(1440, 10, 20));

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertNode(sdtDate);

        doc.save(getArtifactsDir() + "StructuredDocumentTag.Date.docx");
        //ExEnd
    }

    @Test
    public void plainText() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.Color
        //ExFor:StructuredDocumentTag.ContentsFont
        //ExFor:StructuredDocumentTag.EndCharacterFont
        //ExFor:StructuredDocumentTag.Id
        //ExFor:StructuredDocumentTag.Level
        //ExFor:StructuredDocumentTag.Multiline
        //ExFor:StructuredDocumentTag.Tag
        //ExFor:StructuredDocumentTag.Title
        //ExFor:StructuredDocumentTag.RemoveSelfOnly
        //ExSummary:Shows how to create a structured document tag in a plain text box and modify its appearance.
        Document doc = new Document();

        // Create a structured document tag that will contain plain text.
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.INLINE);

        // Set the title and color of the frame that appears when you mouse over the structured document tag in Microsoft Word.
        tag.setTitle("My plain text");
        tag.setColor(Color.MAGENTA);

        // Set a tag for this structured document tag, which is obtainable
        // as an XML element named "tag", with the string below in its "@val" attribute.
        tag.setTag("MyPlainTextSDT");

        // Every structured document tag has a random unique ID.
        Assert.That(tag.getId(), Is.Positive);

        // Set the font for the text inside the structured document tag.
        tag.getContentsFont().setName("Arial");

        // Set the font for the text at the end of the structured document tag.
        // Any text that we type in the document body after moving out of the tag with arrow keys will use this font.
        tag.getEndCharacterFont().setName("Arial Black");

        // By default, this is false and pressing enter while inside a structured document tag does nothing.
        // When set to true, our structured document tag can have multiple lines.

        // Set the "Multiline" property to "false" to only allow the contents
        // of this structured document tag to span a single line.
        // Set the "Multiline" property to "true" to allow the tag to contain multiple lines of content.
        tag.setMultiline(true);

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertNode(tag);

        // Insert a clone of our structured document tag in a new paragraph.
        StructuredDocumentTag tagClone = (StructuredDocumentTag)tag.deepClone(true);
        builder.insertParagraph();
        builder.insertNode(tagClone);

        // Use the "RemoveSelfOnly" method to remove a structured document tag, while keeping its contents in the document.
        tagClone.removeSelfOnly();

        doc.save(getArtifactsDir() + "StructuredDocumentTag.PlainText.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "StructuredDocumentTag.PlainText.docx");
        tag = (StructuredDocumentTag)doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);

        Assert.assertEquals("My plain text", tag.getTitle());
        Assert.assertEquals(Color.MAGENTA.getRGB(), tag.getColor().getRGB());
        Assert.assertEquals("MyPlainTextSDT", tag.getTag());
        Assert.That(tag.getId(), Is.Positive);
        Assert.assertEquals("Arial", tag.getContentsFont().getName());
        Assert.assertEquals("Arial Black", tag.getEndCharacterFont().getName());
        Assert.assertTrue(tag.getMultiline());
    }

    @Test (dataProvider = "isTemporaryDataProvider")
    public void isTemporary(boolean isTemporary) throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.IsTemporary
        //ExSummary:Shows how to make single-use controls.
        Document doc = new Document();

        // Insert a plain text structured document tag,
        // which will act as a plain text form that the user may enter text into.
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.INLINE);

        // Set the "IsTemporary" property to "true" to make the structured document tag disappear and
        // assimilate its contents into the document after the user edits it once in Microsoft Word.
        // Set the "IsTemporary" property to "false" to allow the user to edit the contents
        // of the structured document tag any number of times.
        tag.isTemporary(isTemporary);

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Please enter text: ");
        builder.insertNode(tag);

        // Insert another structured document tag in the form of a check box and set its default state to "checked".
        tag = new StructuredDocumentTag(doc, SdtType.CHECKBOX, MarkupLevel.INLINE);
        tag.setChecked(true);

        // Set the "IsTemporary" property to "true" to make the check box become a symbol
        // once the user clicks on it in Microsoft Word.
        // Set the "IsTemporary" property to "false" to allow the user to click on the check box any number of times.
        tag.isTemporary(isTemporary);

        builder.write("\nPlease click the check box: ");
        builder.insertNode(tag);

        doc.save(getArtifactsDir() + "StructuredDocumentTag.IsTemporary.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "StructuredDocumentTag.IsTemporary.docx");

        Assert.AreEqual(2, 
            doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true).Count(sdt => ((StructuredDocumentTag)sdt).IsTemporary == isTemporary));
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "isTemporaryDataProvider")
	public static Object[][] isTemporaryDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "placeholderBuildingBlockDataProvider")
    public void placeholderBuildingBlock(boolean isShowingPlaceholderText) throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.IsShowingPlaceholderText
        //ExFor:StructuredDocumentTag.Placeholder
        //ExFor:StructuredDocumentTag.PlaceholderName
        //ExSummary:Shows how to use a building block's contents as a custom placeholder text for a structured document tag. 
        Document doc = new Document();

        // Insert a plain text structured document tag of the "PlainText" type, which will function as a text box.
        // The contents that it will display by default are a "Click here to enter text." prompt.
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.INLINE);

        // We can get the tag to display the contents of a building block instead of the default text.
        // First, add a building block with contents to the glossary document.
        GlossaryDocument glossaryDoc = doc.getGlossaryDocument();

        BuildingBlock substituteBlock = new BuildingBlock(glossaryDoc);
        substituteBlock.setName("Custom Placeholder");
        substituteBlock.appendChild(new Section(glossaryDoc));
        substituteBlock.getFirstSection().appendChild(new Body(glossaryDoc));
        substituteBlock.getFirstSection().getBody().appendParagraph("Custom placeholder text.");

        glossaryDoc.appendChild(substituteBlock);

        // Then, use the structured document tag's "PlaceholderName" property to reference that building block by name.
        tag.setPlaceholderName("Custom Placeholder");

        // If "PlaceholderName" refers to an existing block in the parent document's glossary document,
        // we will be able to verify the building block via the "Placeholder" property.
        Assert.assertEquals(substituteBlock, tag.getPlaceholder());

        // Set the "IsShowingPlaceholderText" property to "true" to treat the
        // structured document tag's current contents as placeholder text.
        // This means that clicking on the text box in Microsoft Word will immediately highlight all the tag's contents.
        // Set the "IsShowingPlaceholderText" property to "false" to get the
        // structured document tag to treat its contents as text that a user has already entered.
        // Clicking on this text in Microsoft Word will place the blinking cursor at the clicked location.
        tag.isShowingPlaceholderText(isShowingPlaceholderText);

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertNode(tag);

        doc.save(getArtifactsDir() + "StructuredDocumentTag.PlaceholderBuildingBlock.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "StructuredDocumentTag.PlaceholderBuildingBlock.docx");
        tag = (StructuredDocumentTag)doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        substituteBlock = (BuildingBlock)doc.getGlossaryDocument().getChild(NodeType.BUILDING_BLOCK, 0, true);

        Assert.assertEquals("Custom Placeholder", substituteBlock.getName());
        Assert.assertEquals(isShowingPlaceholderText, tag.isShowingPlaceholderText());
        Assert.assertEquals(substituteBlock, tag.getPlaceholder());
        Assert.assertEquals(substituteBlock.getName(), tag.getPlaceholderName());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "placeholderBuildingBlockDataProvider")
	public static Object[][] placeholderBuildingBlockDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void lock() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.LockContentControl
        //ExFor:StructuredDocumentTag.LockContents
        //ExSummary:Shows how to apply editing restrictions to structured document tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain text structured document tag, which acts as a text box that prompts the user to fill it in.
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.INLINE);

        // Set the "LockContents" property to "true" to prohibit the user from editing this text box's contents.
        tag.setLockContents(true);
        builder.write("The contents of this structured document tag cannot be edited: ");
        builder.insertNode(tag);

        tag = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.INLINE);

        // Set the "LockContentControl" property to "true" to prohibit the user from
        // deleting this structured document tag manually in Microsoft Word.
        tag.setLockContentControl(true);

        builder.insertParagraph();
        builder.write("This structured document tag cannot be deleted but its contents can be edited: ");
        builder.insertNode(tag);

        doc.save(getArtifactsDir() + "StructuredDocumentTag.Lock.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "StructuredDocumentTag.Lock.docx");
        tag = (StructuredDocumentTag)doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);

        Assert.assertTrue(tag.getLockContents());
        Assert.assertFalse(tag.getLockContentControl());

        tag = (StructuredDocumentTag)doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 1, true);

        Assert.assertFalse(tag.getLockContents());
        Assert.assertTrue(tag.getLockContentControl());
    }

    @Test
    public void listItemCollection() throws Exception
    {
        //ExStart
        //ExFor:SdtListItem
        //ExFor:SdtListItem.#ctor(System.String)
        //ExFor:SdtListItem.#ctor(System.String,System.String)
        //ExFor:SdtListItem.DisplayText
        //ExFor:SdtListItem.Value
        //ExFor:SdtListItemCollection
        //ExFor:SdtListItemCollection.Add(Aspose.Words.Markup.SdtListItem)
        //ExFor:SdtListItemCollection.Clear
        //ExFor:SdtListItemCollection.Count
        //ExFor:SdtListItemCollection.GetEnumerator
        //ExFor:SdtListItemCollection.Item(System.Int32)
        //ExFor:SdtListItemCollection.RemoveAt(System.Int32)
        //ExFor:SdtListItemCollection.SelectedValue
        //ExFor:StructuredDocumentTag.ListItems
        //ExSummary:Shows how to work with drop down-list structured document tags.
        Document doc = new Document();
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.DROP_DOWN_LIST, MarkupLevel.BLOCK);
        doc.getFirstSection().getBody().appendChild(tag);

        // A drop-down list structured document tag is a form that allows the user to
        // select an option from a list by left-clicking and opening the form in Microsoft Word.
        // The "ListItems" property contains all list items, and each list item is an "SdtListItem".
        SdtListItemCollection listItems = tag.getListItems();
        listItems.add(new SdtListItem("Value 1"));

        Assert.assertEquals(listItems.get(0).getDisplayText(), listItems.get(0).getValue());

        // Add 3 more list items. Initialize these items using a different constructor to the first item
        // to display strings that are different from their values.
        listItems.add(new SdtListItem("Item 2", "Value 2"));
        listItems.add(new SdtListItem("Item 3", "Value 3"));
        listItems.add(new SdtListItem("Item 4", "Value 4"));

        Assert.assertEquals(4, listItems.getCount());

        // The drop-down list is displaying the first item. Assign a different list item to the "SelectedValue" to display it.
        listItems.setSelectedValue(listItems.get(3));

        Assert.assertEquals("Value 4", listItems.getSelectedValue().getValue());

        // Enumerate over the collection and print each element.
        Iterator<SdtListItem> enumerator = listItems.iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
                if (enumerator.next() != null)
                    System.out.println("List item: {enumerator.Current.DisplayText}, value: {enumerator.Current.Value}");
        }
        finally { if (enumerator != null) enumerator.close(); }

        // Remove the last list item. 
        listItems.removeAt(3);

        Assert.assertEquals(3, listItems.getCount());

        // Since our drop-down control is set to display the removed item by default, give it an item to display which exists.
        listItems.setSelectedValue(listItems.get(1));
       
        doc.save(getArtifactsDir() + "StructuredDocumentTag.ListItemCollection.docx");

        // Use the "Clear" method to empty the entire drop-down item collection at once.
        listItems.clear();

        Assert.assertEquals(0, listItems.getCount());
        //ExEnd
    }

    @Test
    public void creatingCustomXml() throws Exception
    {
        //ExStart
        //ExFor:CustomXmlPart
        //ExFor:CustomXmlPart.Clone
        //ExFor:CustomXmlPart.Data
        //ExFor:CustomXmlPart.Id
        //ExFor:CustomXmlPart.Schemas
        //ExFor:CustomXmlPartCollection
        //ExFor:CustomXmlPartCollection.Add(CustomXmlPart)
        //ExFor:CustomXmlPartCollection.Add(String, String)
        //ExFor:CustomXmlPartCollection.Clear
        //ExFor:CustomXmlPartCollection.Clone
        //ExFor:CustomXmlPartCollection.Count
        //ExFor:CustomXmlPartCollection.GetById(String)
        //ExFor:CustomXmlPartCollection.GetEnumerator
        //ExFor:CustomXmlPartCollection.Item(Int32)
        //ExFor:CustomXmlPartCollection.RemoveAt(Int32)
        //ExFor:Document.CustomXmlParts
        //ExFor:StructuredDocumentTag.XmlMapping
        //ExFor:XmlMapping.SetMapping(CustomXmlPart, String, String)
        //ExSummary:Shows how to create a structured document tag with custom XML data.
        Document doc = new Document();

        // Construct an XML part that contains data and add it to the document's collection.
        // If we enable the "Developer" tab in Microsoft Word,
        // we can find elements from this collection in the "XML Mapping Pane", along with a few default elements.
        String xmlPartId = Guid.newGuid().toString("B");
        String xmlPartContent = "<root><text>Hello world!</text></root>";
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add(xmlPartId, xmlPartContent);

        Assert.assertEquals(Encoding.getASCII().getBytes(xmlPartContent), xmlPart.getData());
        Assert.assertEquals(xmlPartId, xmlPart.getId());

        // Below are two ways to refer to XML parts.
        // 1 -  By an index in the custom XML part collection:
        Assert.assertEquals(xmlPart, doc.getCustomXmlParts().get(0));

        // 2 -  By GUID:
        Assert.assertEquals(xmlPart, doc.getCustomXmlParts().getById(xmlPartId));

        // Add an XML schema association.
        xmlPart.getSchemas().add("http://www.w3.org/2001/XMLSchema");
        
        // Clone a part, and then insert it into the collection.
        CustomXmlPart xmlPartClone = xmlPart.deepClone();
        xmlPartClone.setId(Guid.newGuid().toString("B"));
        doc.getCustomXmlParts().add(xmlPartClone);

        Assert.assertEquals(2, doc.getCustomXmlParts().getCount());

        // Iterate through the collection and print the contents of each part.
        Iterator<CustomXmlPart> enumerator = doc.getCustomXmlParts().iterator();
        try /*JAVA: was using*/
        {
            int index = 0;
            while (enumerator.hasNext())
            {
                System.out.println("XML part index {index}, ID: {enumerator.Current.Id}");
                System.out.println("\tContent: {Encoding.UTF8.GetString(enumerator.Current.Data)}");
                index++;
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // Use the "RemoveAt" method to remove the cloned part by index.
        doc.getCustomXmlParts().removeAt(1);

        Assert.assertEquals(1, doc.getCustomXmlParts().getCount());

        // Clone the XML parts collection, and then use the "Clear" method to remove all its elements at once.
        CustomXmlPartCollection customXmlParts = doc.getCustomXmlParts().deepClone();
        customXmlParts.clear();

        // Create a structured document tag that will display our part's contents and insert it into the document body.
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.BLOCK);
        tag.getXmlMapping().setMapping(xmlPart, "/root[1]/text[1]", "");

        doc.getFirstSection().getBody().appendChild(tag);

        doc.save(getArtifactsDir() + "StructuredDocumentTag.CustomXml.docx");
        //ExEnd

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "StructuredDocumentTag.CustomXml.docx", getGoldsDir() + "StructuredDocumentTag.CustomXml Gold.docx"));

        doc = new Document(getArtifactsDir() + "StructuredDocumentTag.CustomXml.docx");
        xmlPart = doc.getCustomXmlParts().get(0);

        Ref<Guid> referenceToGuid = new Ref<Guid>(Guid);
        Assert.True(Guid.TryParse(xmlPart.getId(), /*out*/ referenceToGuid temp));
        Guid = referenceToGuid.get();
        Assert.assertEquals("<root><text>Hello world!</text></root>", Encoding.getUTF8().getString(xmlPart.getData()));
        Assert.assertEquals("http://www.w3.org/2001/XMLSchema", xmlPart.getSchemas().get(0));

        tag = (StructuredDocumentTag)doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        Assert.assertEquals("Hello world!", tag.getText().trim());
        Assert.assertEquals("/root[1]/text[1]", tag.getXmlMapping().getXPath());
        Assert.assertEquals("", tag.getXmlMapping().getPrefixMappings());
    }

    @Test
    public void xmlMapping() throws Exception
    {
        //ExStart
        //ExFor:XmlMapping
        //ExFor:XmlMapping.CustomXmlPart
        //ExFor:XmlMapping.Delete
        //ExFor:XmlMapping.IsMapped
        //ExFor:XmlMapping.PrefixMappings
        //ExFor:XmlMapping.XPath
        //ExSummary:Shows how to set XML mappings for custom XML parts.
        Document doc = new Document();

        // Construct an XML part that contains text and add it to the document's CustomXmlPart collection.
        String xmlPartId = Guid.newGuid().toString("B");
        String xmlPartContent = "<root><text>Text element #1</text><text>Text element #2</text></root>";
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add(xmlPartId, xmlPartContent);

        Assert.assertEquals("<root><text>Text element #1</text><text>Text element #2</text></root>", 
            Encoding.getUTF8().getString(xmlPart.getData()));

        // Create a structured document tag that will display the contents of our CustomXmlPart.
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.BLOCK);

        // Set a mapping for our structured document tag. This mapping will instruct
        // our structured document tag to display a portion of the XML part's text contents that the XPath points to.
        // In this case, it will be contents of the the second "<text>" element of the first "<root>" element: "Text element #2".
        tag.getXmlMapping().setMapping(xmlPart, "/root[1]/text[2]", "xmlns:ns='http://www.w3.org/2001/XMLSchema'");

        Assert.assertTrue(tag.getXmlMapping().isMapped());
        Assert.assertEquals(xmlPart, tag.getXmlMapping().getCustomXmlPart());
        Assert.assertEquals("/root[1]/text[2]", tag.getXmlMapping().getXPath());
        Assert.assertEquals("xmlns:ns='http://www.w3.org/2001/XMLSchema'", tag.getXmlMapping().getPrefixMappings());

        // Add the structured document tag to the document to display the content from our custom part.
        doc.getFirstSection().getBody().appendChild(tag);
        doc.save(getArtifactsDir() + "StructuredDocumentTag.XmlMapping.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "StructuredDocumentTag.XmlMapping.docx");
        xmlPart = doc.getCustomXmlParts().get(0);

        Ref<Guid> referenceToGuid = new Ref<Guid>(Guid);
        Assert.True(Guid.TryParse(xmlPart.getId(), /*out*/ referenceToGuid temp));
        Guid = referenceToGuid.get();
        Assert.assertEquals("<root><text>Text element #1</text><text>Text element #2</text></root>", Encoding.getUTF8().getString(xmlPart.getData()));

        tag = (StructuredDocumentTag)doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        Assert.assertEquals("Text element #2", tag.getText().trim());
        Assert.assertEquals("/root[1]/text[2]", tag.getXmlMapping().getXPath());
        Assert.assertEquals("xmlns:ns='http://www.w3.org/2001/XMLSchema'", tag.getXmlMapping().getPrefixMappings());
    }

    @Test
    public void structuredDocumentTagRangeStartXmlMapping() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTagRangeStart.XmlMapping
        //ExSummary:Shows how to set XML mappings for the range start of a structured document tag.
        Document doc = new Document(getMyDir() + "Multi-section structured document tags.docx");

        // Construct an XML part that contains text and add it to the document's CustomXmlPart collection.
        String xmlPartId = Guid.newGuid().toString("B");
        String xmlPartContent = "<root><text>Text element #1</text><text>Text element #2</text></root>";
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add(xmlPartId, xmlPartContent);

        Assert.assertEquals("<root><text>Text element #1</text><text>Text element #2</text></root>",
            Encoding.getUTF8().getString(xmlPart.getData()));

        // Create a structured document tag that will display the contents of our CustomXmlPart in the document.
        StructuredDocumentTagRangeStart sdtRangeStart = (StructuredDocumentTagRangeStart)doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG_RANGE_START, 0, true);

        // If we set a mapping for our structured document tag,
        // it will only display a portion of the CustomXmlPart that the XPath points to.
        // This XPath will point to the contents second "<text>" element of the first "<root>" element of our CustomXmlPart.
        sdtRangeStart.getXmlMapping().setMapping(xmlPart, "/root[1]/text[2]", null);

        doc.save(getArtifactsDir() + "StructuredDocumentTag.StructuredDocumentTagRangeStartXmlMapping.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "StructuredDocumentTag.StructuredDocumentTagRangeStartXmlMapping.docx");
        xmlPart = doc.getCustomXmlParts().get(0);

        Ref<Guid> referenceToGuid = new Ref<Guid>(Guid);
        Assert.True(Guid.TryParse(xmlPart.getId(), /*out*/ referenceToGuid temp));
        Guid = referenceToGuid.get();
        Assert.assertEquals("<root><text>Text element #1</text><text>Text element #2</text></root>", Encoding.getUTF8().getString(xmlPart.getData()));

        sdtRangeStart = (StructuredDocumentTagRangeStart)doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG_RANGE_START, 0, true);
        Assert.assertEquals("/root[1]/text[2]", sdtRangeStart.getXmlMapping().getXPath());
    }

    @Test
    public void customXmlSchemaCollection() throws Exception
    {
        //ExStart
        //ExFor:CustomXmlSchemaCollection
        //ExFor:CustomXmlSchemaCollection.Add(System.String)
        //ExFor:CustomXmlSchemaCollection.Clear
        //ExFor:CustomXmlSchemaCollection.Clone
        //ExFor:CustomXmlSchemaCollection.Count
        //ExFor:CustomXmlSchemaCollection.GetEnumerator
        //ExFor:CustomXmlSchemaCollection.IndexOf(System.String)
        //ExFor:CustomXmlSchemaCollection.Item(System.Int32)
        //ExFor:CustomXmlSchemaCollection.Remove(System.String)
        //ExFor:CustomXmlSchemaCollection.RemoveAt(System.Int32)
        //ExSummary:Shows how to work with an XML schema collection.
        Document doc = new Document();

        String xmlPartId = Guid.newGuid().toString("B");
        String xmlPartContent = "<root><text>Hello, World!</text></root>";
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add(xmlPartId, xmlPartContent);

        // Add an XML schema association.
        xmlPart.getSchemas().add("http://www.w3.org/2001/XMLSchema");

        // Clone the custom XML part's XML schema association collection,
        // and then add a couple of new schemas to the clone.
        CustomXmlSchemaCollection schemas = xmlPart.getSchemas().deepClone();
        schemas.add("http://www.w3.org/2001/XMLSchema-instance");
        schemas.add("http://schemas.microsoft.com/office/2006/metadata/contentType");
        
        Assert.assertEquals(3, schemas.getCount());
        Assert.assertEquals(2, schemas.indexOf("http://schemas.microsoft.com/office/2006/metadata/contentType"));

        // Enumerate the schemas and print each element.
        Iterator<String> enumerator = schemas.iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
                System.out.println(enumerator.next());
        }
        finally { if (enumerator != null) enumerator.close(); }

        // Below are three ways of removing schemas from the collection.
        // 1 -  Remove a schema by index:
        schemas.removeAt(2);

        // 2 -  Remove a schema by value:
        schemas.remove("http://www.w3.org/2001/XMLSchema");

        // 3 -  Use the "Clear" method to empty the collection at once.
        schemas.clear();

        Assert.assertEquals(0, schemas.getCount());
        //ExEnd
    }

    @Test
    public void customXmlPartStoreItemIdReadOnly() throws Exception
    {
        //ExStart
        //ExFor:XmlMapping.StoreItemId
        //ExSummary:Shows how to get the custom XML data identifier of an XML part.
        Document doc = new Document(getMyDir() + "Custom XML part in structured document tag.docx");

        // Structured document tags have IDs in the form of GUIDs.
        StructuredDocumentTag tag = (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);

        Assert.assertEquals("{F3029283-4FF8-4DD2-9F31-395F19ACEE85}", tag.getXmlMapping().getStoreItemId());
        //ExEnd
    }

    @Test
    public void customXmlPartStoreItemIdReadOnlyNull() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        StructuredDocumentTag sdtCheckBox = 
            new StructuredDocumentTag(doc, SdtType.CHECKBOX, MarkupLevel.INLINE); {sdtCheckBox.setChecked(true);}

        builder.insertNode(sdtCheckBox);

        doc = DocumentHelper.saveOpen(doc);

        StructuredDocumentTag sdt = (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        System.out.println("The Id of your custom xml part is: " + sdt.getXmlMapping().getStoreItemId());
    }

    @Test
    public void clearTextFromStructuredDocumentTags() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.Clear
        //ExSummary:Shows how to delete contents of structured document tag elements.
        Document doc = new Document();

        // Create a plain text structured document tag, and then append it to the document.
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.BLOCK);
        doc.getFirstSection().getBody().appendChild(tag);

        // This structured document tag, which is in the form of a text box, already displays placeholder text.
        Assert.assertEquals("Click here to enter text.", tag.getText().trim());
        Assert.assertTrue(tag.isShowingPlaceholderText());

        // Create a building block with text contents.
        GlossaryDocument glossaryDoc = doc.getGlossaryDocument();
        BuildingBlock substituteBlock = new BuildingBlock(glossaryDoc);
        substituteBlock.setName("My placeholder");
        substituteBlock.appendChild(new Section(glossaryDoc));
        substituteBlock.getFirstSection().ensureMinimum();
        substituteBlock.getFirstSection().getBody().getFirstParagraph().appendChild(new Run(glossaryDoc, "Custom placeholder text."));
        glossaryDoc.appendChild(substituteBlock);

        // Set the structured document tag's "PlaceholderName" property to our building block's name to get
        // the structured document tag to display the contents of the building block in place of the original default text.
        tag.setPlaceholderName("My placeholder");

        Assert.assertEquals("Custom placeholder text.", tag.getText().trim());
        Assert.assertTrue(tag.isShowingPlaceholderText());

        // Edit the text of the structured document tag and hide the placeholder text.
        Run run = (Run)tag.getChild(NodeType.RUN, 0, true);
        run.setText("New text.");
        tag.isShowingPlaceholderText(false);

        Assert.assertEquals("New text.", tag.getText().trim());

        // Use the "Clear" method to clear this structured document tag's contents and display the placeholder again.
        tag.clear();

        Assert.assertTrue(tag.isShowingPlaceholderText());
        Assert.assertEquals("Custom placeholder text.", tag.getText().trim());
        //ExEnd
    }

    @Test
    public void accessToBuildingBlockPropertiesFromDocPartObjSdt() throws Exception
    {
        Document doc = new Document(getMyDir() + "Structured document tags with building blocks.docx");

        StructuredDocumentTag docPartObjSdt =
            (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);

        Assert.assertEquals(SdtType.DOC_PART_OBJ, docPartObjSdt.getSdtType());
        Assert.assertEquals("Table of Contents", docPartObjSdt.getBuildingBlockGallery());
    }

    @Test
    public void accessToBuildingBlockPropertiesFromPlainTextSdt() throws Exception
    {
        Document doc = new Document(getMyDir() + "Structured document tags with building blocks.docx");

        StructuredDocumentTag plainTextSdt =
            (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 1, true);

        Assert.assertEquals(SdtType.PLAIN_TEXT, plainTextSdt.getSdtType());
        Assert.That(() => plainTextSdt.getBuildingBlockGallery(), Throws.<IllegalStateException>TypeOf(),
            "BuildingBlockType is only accessible for BuildingBlockGallery SDT type.");
    }

    @Test
    public void buildingBlockCategories() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.BuildingBlockCategory
        //ExFor:StructuredDocumentTag.BuildingBlockGallery
        //ExSummary:Shows how to insert a structured document tag as a building block, and set its category and gallery.
        Document doc = new Document();

        StructuredDocumentTag buildingBlockSdt =
            new StructuredDocumentTag(doc, SdtType.BUILDING_BLOCK_GALLERY, MarkupLevel.BLOCK);
            {
                buildingBlockSdt.setBuildingBlockCategory("Built-in");
                buildingBlockSdt.setBuildingBlockGallery("Table of Contents");
            }

        doc.getFirstSection().getBody().appendChild(buildingBlockSdt);

        doc.save(getArtifactsDir() + "StructuredDocumentTag.BuildingBlockCategories.docx");
        //ExEnd

        buildingBlockSdt =
            (StructuredDocumentTag) doc.getFirstSection().getBody().getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);

        Assert.assertEquals(SdtType.BUILDING_BLOCK_GALLERY, buildingBlockSdt.getSdtType());
        Assert.assertEquals("Table of Contents", buildingBlockSdt.getBuildingBlockGallery());
        Assert.assertEquals("Built-in", buildingBlockSdt.getBuildingBlockCategory());
    }

    @Test (dataProvider = "updateSdtContentDataProvider")
    public void updateSdtContent(boolean updateSdtContent) throws Exception
    {
        //ExStart
        //ExFor:SaveOptions.UpdateSdtContent
        //ExSummary:Shows how to update structured document tags while saving a document to PDF.
        Document doc = new Document();

        // Insert a drop-down list structured document tag.
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.DROP_DOWN_LIST, MarkupLevel.BLOCK);
        tag.getListItems().add(new SdtListItem("Value 1"));
        tag.getListItems().add(new SdtListItem("Value 2"));
        tag.getListItems().add(new SdtListItem("Value 3"));

        // The drop-down list currently displays "Choose an item" as the default text.
        // Set the "SelectedValue" property to one of the list items to get the tag to
        // display that list item's value instead of the default text.
        tag.getListItems().setSelectedValue(tag.getListItems().get(1));

        doc.getFirstSection().getBody().appendChild(tag);

        // Create a "PdfSaveOptions" object to pass to the document's "Save" method
        // to modify how that method saves the document to .PDF.
        PdfSaveOptions options = new PdfSaveOptions();

        // Set the "UpdateSdtContent" property to "false" not to update the structured document tags
        // while saving the document to PDF. They will display their default values as they were at the time of construction.
        // Set the "UpdateSdtContent" property to "true" to make sure the tags display updated values in the PDF.
        options.setUpdateSdtContent(updateSdtContent);

        doc.save(getArtifactsDir() + "StructuredDocumentTag.UpdateSdtContent.pdf", options);
        //ExEnd

        Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(getArtifactsDir() + "StructuredDocumentTag.UpdateSdtContent.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();
        textAbsorber.Visit(pdfDoc);

        Assert.AreEqual(updateSdtContent ? "Value 2" : "Choose an item.",
            textAbsorber.Text);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "updateSdtContentDataProvider")
	public static Object[][] updateSdtContentDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void fillTableUsingRepeatingSectionItem() throws Exception
    {
        //ExStart
        //ExFor:SdtType
        //ExSummary:Shows how to fill a table with data from in an XML part.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
 
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add("Books",
            "<books>" +
                "<book>" +
                    "<title>Everyday Italian</title>" +
                    "<author>Giada De Laurentiis</author>" +
                "</book>" +
                "<book>" +
                    "<title>The C Programming Language</title>" +
                    "<author>Brian W. Kernighan, Dennis M. Ritchie</author>" +
                "</book>" +
                "<book>" +
                    "<title>Learning XML</title>" +
                    "<author>Erik T. Ray</author>" +
                "</book>" +
            "</books>");
 
        // Create headers for data from the XML content.
        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Title");
        builder.insertCell();
        builder.write("Author");
        builder.endRow();
        builder.endTable();

        // Create a table with a repeating section inside.
        StructuredDocumentTag repeatingSectionSdt =
            new StructuredDocumentTag(doc, SdtType.REPEATING_SECTION, MarkupLevel.ROW);
        repeatingSectionSdt.getXmlMapping().setMapping(xmlPart, "/books[1]/book", "");
        table.appendChild(repeatingSectionSdt);

        // Add repeating section item inside the repeating section and mark it as a row.
        // This table will have a row for each element that we can find in the XML document
        // using the "/books[1]/book" XPath, of which there are three.
        StructuredDocumentTag repeatingSectionItemSdt =
            new StructuredDocumentTag(doc, SdtType.REPEATING_SECTION_ITEM, MarkupLevel.ROW);
        repeatingSectionSdt.appendChild(repeatingSectionItemSdt);
 
        Row row = new Row(doc);
        repeatingSectionItemSdt.appendChild(row);

        // Map XML data with created table cells for the title and author of each book.
        StructuredDocumentTag titleSdt =
            new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.CELL);
        titleSdt.getXmlMapping().setMapping(xmlPart, "/books[1]/book[1]/title[1]", "");
        row.appendChild(titleSdt);
 
        StructuredDocumentTag authorSdt =
            new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.CELL);
        authorSdt.getXmlMapping().setMapping(xmlPart, "/books[1]/book[1]/author[1]", "");
        row.appendChild(authorSdt);
 
        doc.save(getArtifactsDir() + "StructuredDocumentTag.RepeatingSectionItem.docx");
		//ExEnd

        doc = new Document(getArtifactsDir() + "StructuredDocumentTag.RepeatingSectionItem.docx");
        ArrayList<StructuredDocumentTag> tags = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true).<StructuredDocumentTag>OfType().ToList();

        Assert.assertEquals("/books[1]/book", tags.get(0).getXmlMapping().getXPath());
        Assert.assertEquals("", tags.get(0).getXmlMapping().getPrefixMappings());

        Assert.assertEquals("", tags.get(1).getXmlMapping().getXPath());
        Assert.assertEquals("", tags.get(1).getXmlMapping().getPrefixMappings());

        Assert.assertEquals("/books[1]/book[1]/title[1]", tags.get(2).getXmlMapping().getXPath());
        Assert.assertEquals("", tags.get(2).getXmlMapping().getPrefixMappings());

        Assert.assertEquals("/books[1]/book[1]/author[1]", tags.get(3).getXmlMapping().getXPath());
        Assert.assertEquals("", tags.get(3).getXmlMapping().getPrefixMappings());

        Assert.assertEquals("Title\u0007Author\u0007\u0007" +
                        "Everyday Italian\u0007Giada De Laurentiis\u0007\u0007" +
                        "The C Programming Language\u0007Brian W. Kernighan, Dennis M. Ritchie\u0007\u0007" +
                        "Learning XML\u0007Erik T. Ray\u0007\u0007", doc.getFirstSection().getBody().getTables().get(0).getText().trim());
    }

    @Test
    public void customXmlPart() throws Exception
    {
        String xmlString = 
           "<?xml version=\"1.0\"?>" +
           "<Company>" +
               "<Employee id=\"1\">" +
                   "<FirstName>John</FirstName>" +
                   "<LastName>Doe</LastName>" +
               "</Employee>" +
               "<Employee id=\"2\">" +
                   "<FirstName>Jane</FirstName>" +
                   "<LastName>Doe</LastName>" +
               "</Employee>" +
           "</Company>";

        Document doc = new Document();

        // Insert the full XML document as a custom document part.
        // We can find the mapping for this part in Microsoft Word via "Developer" -> "XML Mapping Pane", if it is enabled.
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add(Guid.newGuid().toString("B"), xmlString);

        // Create a structured document tag, which will use an XPath to refer to a single element from the XML.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.BLOCK);
        sdt.getXmlMapping().setMapping(xmlPart, "Company//Employee[@id='2']/FirstName", "");

        // Add the StructuredDocumentTag to the document to display the element in the text.
        doc.getFirstSection().getBody().appendChild(sdt);
    }

    @Test
    public void multiSectionTags() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTagRangeStart
        //ExFor:StructuredDocumentTagRangeStart.Id
        //ExFor:StructuredDocumentTagRangeStart.Title
        //ExFor:StructuredDocumentTagRangeStart.IsShowingPlaceholderText
        //ExFor:StructuredDocumentTagRangeStart.LockContentControl
        //ExFor:StructuredDocumentTagRangeStart.LockContents
        //ExFor:StructuredDocumentTagRangeStart.Level
        //ExFor:StructuredDocumentTagRangeStart.RangeEnd
        //ExFor:StructuredDocumentTagRangeStart.SdtType
        //ExFor:StructuredDocumentTagRangeStart.Tag
        //ExFor:StructuredDocumentTagRangeEnd
        //ExFor:StructuredDocumentTagRangeEnd.Id
        //ExSummary:Shows how to get the properties of multi-section structured document tags.
        Document doc = new Document(getMyDir() + "Multi-section structured document tags.docx");

        StructuredDocumentTagRangeStart rangeStartTag =
            ms.as(doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG_RANGE_START, true).get(0), StructuredDocumentTagRangeStart.class);
        StructuredDocumentTagRangeEnd rangeEndTag =
            ms.as(doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG_RANGE_END, true).get(0), StructuredDocumentTagRangeEnd.class);

        Assert.assertEquals(rangeStartTag.getId(), rangeEndTag.getId()); //ExSkip
        Assert.assertEquals(NodeType.STRUCTURED_DOCUMENT_TAG_RANGE_START, rangeStartTag.getNodeType()); //ExSkip
        Assert.assertEquals(NodeType.STRUCTURED_DOCUMENT_TAG_RANGE_END, rangeEndTag.getNodeType()); //ExSkip

        System.out.println("StructuredDocumentTagRangeStart values:");
        System.out.println("\t|Id: {rangeStartTag.Id}");
        System.out.println("\t|Title: {rangeStartTag.Title}");
        System.out.println("\t|IsShowingPlaceholderText: {rangeStartTag.IsShowingPlaceholderText}");
        System.out.println("\t|LockContentControl: {rangeStartTag.LockContentControl}");
        System.out.println("\t|LockContents: {rangeStartTag.LockContents}");
        System.out.println("\t|Level: {rangeStartTag.Level}");
        System.out.println("\t|NodeType: {rangeStartTag.NodeType}");
        System.out.println("\t|RangeEnd: {rangeStartTag.RangeEnd}");
        System.out.println("\t|SdtType: {rangeStartTag.SdtType}");
        System.out.println("\t|Tag: {rangeStartTag.Tag}\n");

        System.out.println("StructuredDocumentTagRangeEnd values:");
        System.out.println("\t|Id: {rangeEndTag.Id}");
        System.out.println("\t|NodeType: {rangeEndTag.NodeType}");
        //ExEnd
    }

    @Test
    public void sdtChildNodes() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTagRangeStart.ChildNodes
        //ExFor:StructuredDocumentTagRangeStart.GetChildNodes(NodeType, bool)
        //ExSummary:Shows how to get child nodes of StructuredDocumentTagRangeStart.
        Document doc = new Document(getMyDir() + "Multi-section structured document tags.docx");
        StructuredDocumentTagRangeStart tag = ms.as(doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG_RANGE_START, true).get(0), StructuredDocumentTagRangeStart.class);

        System.out.println("StructuredDocumentTagRangeStart values:");
        System.out.println("\t|Child nodes count: {tag.ChildNodes.Count}\n");

        for (Node node : (Iterable<Node>) tag.getChildNodes())
            System.out.println("\t|Child node type: {node.NodeType}");

        for (Node node : (Iterable<Node>) tag.getChildNodes(NodeType.RUN, true))
            System.out.println("\t|Child node text: {node.GetText()}");
        //ExEnd
    }
}
