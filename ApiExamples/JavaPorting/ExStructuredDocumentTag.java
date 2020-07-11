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
import com.aspose.ms.System.msString;
import com.aspose.words.CustomXmlSchemaCollection;
import com.aspose.words.Run;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.Table;
import com.aspose.words.Row;
import com.aspose.words.ref.Ref;
import org.testng.annotations.DataProvider;


/// <summary>
/// Tests that verify work with structured document tags in the document. 
/// </summary>
@Test
class ExStructuredDocumentTag !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void repeatingSection() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.SdtType
        //ExSummary:Shows how to get type of structured document tag.
        Document doc = new Document(getMyDir() + "Structured document tags.docx");

        ArrayList<StructuredDocumentTag> sdTags = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true).<StructuredDocumentTag>OfType().ToList();

        Assert.assertEquals(SdtType.REPEATING_SECTION, sdTags.get(0).getSdtType());
        Assert.assertEquals(SdtType.REPEATING_SECTION_ITEM, sdTags.get(1).getSdtType());
        Assert.assertEquals(SdtType.RICH_TEXT, sdTags.get(2).getSdtType());
        //ExEnd
    }

    @Test
    public void setSpecificStyleToSdt() throws Exception
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

        // Get specific style from the document to apply it to an SDT
        Style quoteStyle = doc.getStyles().getByStyleIdentifier(StyleIdentifier.QUOTE);
        StructuredDocumentTag sdtPlainText = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.INLINE);
        sdtPlainText.setStyle(quoteStyle);

        StructuredDocumentTag sdtRichText = new StructuredDocumentTag(doc, SdtType.RICH_TEXT, MarkupLevel.INLINE);
        // Second method to apply specific style to an SDT control
        sdtRichText.setStyleName("Quote");

        // Insert content controls into the document
        builder.insertNode(sdtPlainText);
        builder.insertNode(sdtRichText);

        // We can get a collection of StructuredDocumentTags by looking for the document's child nodes of this NodeType
        Assert.assertEquals(NodeType.STRUCTURED_DOCUMENT_TAG, sdtPlainText.getNodeType());

        NodeCollection tags = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        for (Node node : (Iterable<Node>) tags)
        {
            StructuredDocumentTag sdt = (StructuredDocumentTag)node;
            // If style was not defined before, style should be "Default Paragraph Font"
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
        //ExSummary:Show how to create and insert checkbox structured document tag.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        StructuredDocumentTag sdtCheckBox = new StructuredDocumentTag(doc, SdtType.CHECKBOX, MarkupLevel.INLINE);
        sdtCheckBox.setChecked(true);

        // Insert content control into the document
        builder.insertNode(sdtCheckBox);
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        NodeCollection sdts = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        StructuredDocumentTag sdt = (StructuredDocumentTag) sdts.get(0);
        Assert.assertEquals(true, sdt.getChecked());
        Assert.That(sdt.getXmlMapping().getStoreItemId(), Is.Empty); //Assert that this sdt has no StoreItemId
    }

    @Test
    public void date() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.CalendarType
        //ExFor:StructuredDocumentTag.DateDisplayFormat
        //ExFor:StructuredDocumentTag.DateDisplayLocale
        //ExFor:StructuredDocumentTag.DateStorageFormat
        //ExFor:StructuredDocumentTag.FullDate
        //ExSummary:Shows how to prompt the user to enter a date with a StructuredDocumentTag.
        // Create a new document
        Document doc = new Document();

        // Insert a StructuredDocumentTag that prompts the user to enter a date
        // In Microsoft Word, this element is known as a "Date picker content control"
        // When we click on the arrow on the right end of this tag in Microsoft Word,
        // we will see a pop up in the form of a clickable calendar
        // We can use that popup to select a date that will be displayed by the tag 
        StructuredDocumentTag sdtDate = new StructuredDocumentTag(doc, SdtType.DATE, MarkupLevel.INLINE);

        // This attribute sets the language that the calendar will be displayed in,
        // which in this case will be Saudi Arabian Arabic
        sdtDate.setDateDisplayLocale(msCultureInfo.getCultureInfo("ar-SA").getLCID());

        // We can set the format with which to display the date like this
        // The locale we set above will be carried over to the displayed date
        sdtDate.setDateDisplayFormat("dd MMMM, yyyy");

        // Select how the data will be stored in the document
        sdtDate.setDateStorageFormat(SdtDateStorageFormat.DATE_TIME);

        // Set the calendar type that will be used to select and display the date
        sdtDate.setCalendarType(SdtCalendarType.HIJRI);

        // Before a date is chosen, the tag will display the text "Click here to enter a date."
        // We can set a default date to display by setting this variable
        // We must convert the date to the appropriate calendar ourselves
        sdtDate.setFullDateInternal(new DateTime(1440, 10, 20));

        // Insert the StructuredDocumentTag into the document with a DocumentBuilder and save the document
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
        //ExSummary:Shows how to create a StructuredDocumentTag in the form of a plain text box and modify its appearance.
        // Create a new document 
        Document doc = new Document();

        // Create a StructuredDocumentTag that will contain plain text
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.INLINE);

        // Set the title and color of the frame that appears when you mouse over it
        tag.setTitle("My plain text");
        tag.setColor(Color.MAGENTA);

        // Set a programmatic tag for this StructuredDocumentTag
        // Unlike the title, this value will not be visible in the document but will be programmatically obtainable
        // as an XML element named "tag", with the string below in its "@val" attribute
        tag.setTag("MyPlainTextSDT");

        // Every StructuredDocumentTag gets a random unique ID
        Assert.That(tag.getId(), Is.Positive);

        // Set the font for the text inside the StructuredDocumentTag
        tag.getContentsFont().setName("Arial");

        // Set the font for the text at the end of the StructuredDocumentTag
        // Any text that's typed in the document body after moving out of the tag with arrow keys will keep this font
        tag.getEndCharacterFont().setName("Arial Black");

        // By default, this is false and pressing enter while inside a StructuredDocumentTag does nothing
        // When set to true, our StructuredDocumentTag can have multiple lines
        tag.setMultiline(true);

        // Insert the StructuredDocumentTag into the document with a DocumentBuilder and save the document to a file
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertNode(tag);

        // Insert a clone of our StructuredDocumentTag in a new paragraph
        StructuredDocumentTag tagClone = (StructuredDocumentTag)tag.deepClone(true);
        builder.insertParagraph();
        builder.insertNode(tagClone);

        // We can remove the tag while keeping its contents where they were in the Paragraph by calling RemoveSelfOnly()
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

    @Test
    public void isTemporary() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.IsTemporary
        //ExSummary:Demonstrates the effects of making a StructuredDocumentTag temporary.
        Document doc = new Document();

        // Insert a plain text StructuredDocumentTag, which will prompt the user to enter text
        // and allow them to edit it like a text box
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.INLINE);

        // If we set its Temporary attribute to true, as soon as we start typing,
        // the tag will disappear and its contents will be assimilated into the parent Paragraph
        tag.isTemporary(true);

        // Insert the StructuredDocumentTag with a DocumentBuilder
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Temporary text box: ");
        builder.insertNode(tag);

        // A StructuredDocumentTag in the form of a check box will let the user a square to check and uncheck
        // Setting it to temporary will freeze its value after the first time it is clicked
        tag = new StructuredDocumentTag(doc, SdtType.CHECKBOX, MarkupLevel.INLINE);
        tag.isTemporary(true);

        builder.write("\nTemporary checkbox: ");
        builder.insertNode(tag);

        doc.save(getArtifactsDir() + "StructuredDocumentTag.IsTemporary.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "StructuredDocumentTag.IsTemporary.docx");

        Assert.AreEqual(2, doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true).Count(sdt => ((StructuredDocumentTag)sdt).IsTemporary));
    }

    @Test
    public void placeholderBuildingBlock() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.IsShowingPlaceholderText
        //ExFor:StructuredDocumentTag.Placeholder
        //ExFor:StructuredDocumentTag.PlaceholderName
        //ExSummary:Shows how to use the contents of a BuildingBlock as a custom placeholder text for a StructuredDocumentTag. 
        Document doc = new Document();

        // Insert a plain text StructuredDocumentTag of the PlainText type, which will function like a text box
        // It contains a default "Click here to enter text." prompt, which we can click and replace with our own text
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.INLINE);

        // We can substitute that default placeholder with a custom phrase, which will be drawn from a BuildingBlock
        // First we will need to create the BuildingBlock, give it content and add it to the GlossaryDocument
        GlossaryDocument glossaryDoc = doc.getGlossaryDocument();

        BuildingBlock substituteBlock = new BuildingBlock(glossaryDoc);
        substituteBlock.setName("Custom Placeholder");
        substituteBlock.appendChild(new Section(glossaryDoc));
        substituteBlock.getFirstSection().appendChild(new Body(glossaryDoc));
        substituteBlock.getFirstSection().getBody().appendParagraph("Custom placeholder text.");

        glossaryDoc.appendChild(substituteBlock);

        // The substitute BuildingBlock we made can be referenced by name
        tag.setPlaceholderName("Custom Placeholder");

        // If PlaceholderName refers to an existing block in the parent document's GlossaryDocument,
        // the BuildingBlock will be automatically found and assigned to the Placeholder attribute
        Assert.assertEquals(substituteBlock, tag.getPlaceholder());

        // Setting this to true will register the text inside the StructuredDocumentTag as placeholder text
        // This means that, in Microsoft Word, all the text contents of the StructuredDocumentTag will be highlighted with one click,
        // so we can immediately replace the entire substitute text by typing
        // If this is false, the text will behave like an ordinary Paragraph and a cursor will be placed with nothing highlighted
        tag.isShowingPlaceholderText(true);

        // Insert the StructuredDocumentTag into the document using a DocumentBuilder and save the document to a file
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertNode(tag);

        doc.save(getArtifactsDir() + "StructuredDocumentTag.PlaceholderBuildingBlock.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "StructuredDocumentTag.PlaceholderBuildingBlock.docx");
        tag = (StructuredDocumentTag)doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        substituteBlock = (BuildingBlock)doc.getGlossaryDocument().getChild(NodeType.BUILDING_BLOCK, 0, true);

        Assert.assertEquals("Custom Placeholder", substituteBlock.getName());
        Assert.assertTrue(tag.isShowingPlaceholderText());
        Assert.assertEquals(substituteBlock, tag.getPlaceholder());
        Assert.assertEquals(substituteBlock.getName(), tag.getPlaceholderName());
    }

    @Test
    public void lock() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.LockContentControl
        //ExFor:StructuredDocumentTag.LockContents
        //ExSummary:Shows how to restrict the editing of a StructuredDocumentTag.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain text StructuredDocumentTag of the PlainText type, which will function like a text box
        // It contains a default "Click here to enter text." prompt, which we can click and replace with our own text
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.INLINE);

        // We can prohibit the users from editing the inner text in Microsoft Word by setting this to true
        tag.setLockContents(true);
        builder.write("The contents of this StructuredDocumentTag cannot be edited: ");
        builder.insertNode(tag);

        tag = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.INLINE);

        // Setting this to true will disable the deletion of this StructuredDocumentTag
        // by text editing operations in Microsoft Word
        tag.setLockContentControl(true);

        builder.insertParagraph();
        builder.write("This StructuredDocumentTag cannot be deleted but its contents can be edited: ");
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
        //ExSummary:Shows how to work with StructuredDocumentTag nodes of the DropDownList type.
        // Create a blank document and insert a StructuredDocumentTag that will contain a drop down list
        Document doc = new Document();
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.DROP_DOWN_LIST, MarkupLevel.BLOCK);
        doc.getFirstSection().getBody().appendChild(tag);

        // A drop down list needs elements, each of which will be a SdtListItem
        SdtListItemCollection listItems = tag.getListItems();
        listItems.add(new SdtListItem("Value 1"));

        // Each SdtListItem has text that will be displayed when the drop down list is opened, and also a value
        // When we initialize with one string, we are providing just the value
        // Accordingly, value is passed as DisplayText and will consequently be displayed on the screen
        Assert.assertEquals(listItems.get(0).getDisplayText(), listItems.get(0).getValue());

        // Add 3 more SdtListItems with non-empty strings passed to DisplayText
        listItems.add(new SdtListItem("Item 2", "Value 2"));
        listItems.add(new SdtListItem("Item 3", "Value 3"));
        listItems.add(new SdtListItem("Item 4", "Value 4"));

        // We can obtain a count of the SdtListItems and also set the drop down list's SelectedValue attribute to
        // automatically have one of them pre-selected when we open the document in Microsoft Word
        Assert.assertEquals(4, listItems.getCount());
        listItems.setSelectedValue(listItems.get(3));

        Assert.assertEquals("Value 4", listItems.getSelectedValue().getValue());

        // We can enumerate over the collection and print each element
        Iterator<SdtListItem> enumerator = listItems.iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                if (enumerator.next() != null)
                    System.out.println("List item: {enumerator.Current.DisplayText}, value: {enumerator.Current.Value}");
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // We can also remove elements one at a time
        listItems.removeAt(3);
        Assert.assertEquals(3, listItems.getCount());

        // Make sure to update the SelectedValue's index if it ever ends up out of bounds before saving the document
        listItems.setSelectedValue(listItems.get(1));
       
        doc.save(getArtifactsDir() + "StructuredDocumentTag.ListItemCollection.docx");

        // We can clear the whole collection at once too
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
        //ExSummary:Shows how to create structured document tag with a custom XML data.
        Document doc = new Document();

        // Construct an XML part that contains data and add it to the document's collection
        // Once the "Developer" tab in Mircosoft Word is enabled,
        // we can find elements from this collection as well as a couple defaults in the "XML Mapping Pane" 
        String xmlPartId = Guid.newGuid().toString("B");
        String xmlPartContent = "<root><text>Hello world!</text></root>";
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add(xmlPartId, xmlPartContent);

        // The data we entered is stored in these attributes
        Assert.assertEquals(Encoding.getASCII().getBytes(xmlPartContent), xmlPart.getData());
        Assert.assertEquals(xmlPartId, xmlPart.getId());

        // XML parts can be referenced by collection index or GUID
        Assert.assertEquals(xmlPart, doc.getCustomXmlParts().get(0));
        Assert.assertEquals(xmlPart, doc.getCustomXmlParts().getById(xmlPartId));

        // Once the part is created, we can add XML schema associations like this
        xmlPart.getSchemas().add("http://www.w3.org/2001/XMLSchema");
        
        // We can also clone parts and insert them into the collection directly
        CustomXmlPart xmlPartClone = xmlPart.deepClone();
        xmlPartClone.setId(Guid.newGuid().toString("B"));
        doc.getCustomXmlParts().add(xmlPartClone);

        Assert.assertEquals(2, doc.getCustomXmlParts().getCount());

        // Iterate through collection with an enumerator and print the contents of each part
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

        // XML parts can be removed by index
        doc.getCustomXmlParts().removeAt(1);

        Assert.assertEquals(1, doc.getCustomXmlParts().getCount());

        // The XML part collection itself can be cloned also
        CustomXmlPartCollection customXmlParts = doc.getCustomXmlParts().deepClone();

        // And all elements can be cleared like this
        customXmlParts.clear();

        // Create a StructuredDocumentTag that will display the contents of our part,
        // insert it into the document and save the document
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
        Assert.assertEquals("Hello world!", msString.trim(tag.getText()));
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
        //ExSummary:Shows how to set XML mappings for CustomXmlParts.
        Document doc = new Document();

        // Construct an XML part that contains data and add it to the document's CustomXmlPart collection
        String xmlPartId = Guid.newGuid().toString("B");
        String xmlPartContent = "<root><text>Text element #1</text><text>Text element #2</text></root>";
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add(xmlPartId, xmlPartContent);
        System.out.println(Encoding.getUTF8().getString(xmlPart.getData()));

        // Create a StructuredDocumentTag that will display the contents of our CustomXmlPart in the document
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.BLOCK);

        // If we set a mapping for our StructuredDocumentTag,
        // it will only display a part of the CustomXmlPart that the XPath points to
        // This XPath will point to the contents second "<text>" element of the first "<root>" element of our CustomXmlPart
        tag.getXmlMapping().setMapping(xmlPart, "/root[1]/text[2]", "xmlns:ns='http://www.w3.org/2001/XMLSchema'");

        Assert.assertTrue(tag.getXmlMapping().isMapped());
        Assert.assertEquals(xmlPart, tag.getXmlMapping().getCustomXmlPart());
        Assert.assertEquals("/root[1]/text[2]", tag.getXmlMapping().getXPath());
        Assert.assertEquals("xmlns:ns='http://www.w3.org/2001/XMLSchema'", tag.getXmlMapping().getPrefixMappings());

        // Add the StructuredDocumentTag to the document to display the content from our CustomXmlPart
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
        Assert.assertEquals("Text element #2", msString.trim(tag.getText()));
        Assert.assertEquals("/root[1]/text[2]", tag.getXmlMapping().getXPath());
        Assert.assertEquals("xmlns:ns='http://www.w3.org/2001/XMLSchema'", tag.getXmlMapping().getPrefixMappings());
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
        // Create a document and add a custom XML part
        Document doc = new Document();

        String xmlPartId = Guid.newGuid().toString("B");
        String xmlPartContent = "<root><text>Hello, World!</text></root>";
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add(xmlPartId, xmlPartContent);

        // Once the part is created, we can add XML schema associations like this,
        // and perform other collection-related operations on the list of schemas for this part
        xmlPart.getSchemas().add("http://www.w3.org/2001/XMLSchema");

        // Collections can be cloned and elements can be added
        CustomXmlSchemaCollection schemas = xmlPart.getSchemas().deepClone();
        schemas.add("http://www.w3.org/2001/XMLSchema-instance");
        schemas.add("http://schemas.microsoft.com/office/2006/metadata/contentType");
        
        Assert.assertEquals(3, schemas.getCount());
        Assert.assertEquals(2, schemas.indexOf(("http://schemas.microsoft.com/office/2006/metadata/contentType")));

        // We can iterate over the collection with an enumerator
        Iterator<String> enumerator = schemas.iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                System.out.println(enumerator.next());
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // We can also remove elements by index, element, or we can clear the entire collection
        schemas.removeAt(2);
        schemas.remove("http://www.w3.org/2001/XMLSchema");
        schemas.clear();

        Assert.assertEquals(0, schemas.getCount());
        //ExEnd
    }

    @Test
    public void customXmlPartStoreItemIdReadOnly() throws Exception
    {
        //ExStart
        //ExFor:XmlMapping.StoreItemId
        //ExSummary:Shows how to get special id of your xml part.
        Document doc = new Document(getMyDir() + "Custom XML part in structured document tag.docx");

        // Structured document tags have IDs in the form of Guids
        StructuredDocumentTag tag = (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        Assert.assertEquals("{F3029283-4FF8-4DD2-9F31-395F19ACEE85}", tag.getXmlMapping().getStoreItemId());
        //ExEnd
    }

    @Test
    public void customXmlPartStoreItemIdReadOnlyNull() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        StructuredDocumentTag sdtCheckBox = new StructuredDocumentTag(doc, SdtType.CHECKBOX, MarkupLevel.INLINE);
        sdtCheckBox.setChecked(true);

        // Insert content control into the document
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
        //ExSummary:Shows how to delete content of StructuredDocumentTag elements.
        Document doc = new Document();

        // Create a plain text structured document tag and append it to the document
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.BLOCK);
        doc.getFirstSection().getBody().appendChild(tag);

        // This structured document tag, which is in the form of a text box, already displays placeholder text
        Assert.assertEquals("Click here to enter text.", msString.trim(tag.getText()));
        Assert.assertTrue(tag.isShowingPlaceholderText());

        // Create a building block that 
        GlossaryDocument glossaryDoc = doc.getGlossaryDocument();
        BuildingBlock substituteBlock = new BuildingBlock(glossaryDoc);
        substituteBlock.setName("My placeholder");
        substituteBlock.appendChild(new Section(glossaryDoc));
        substituteBlock.getFirstSection().ensureMinimum();
        substituteBlock.getFirstSection().getBody().getFirstParagraph().appendChild(new Run(glossaryDoc, "Custom placeholder text."));
        glossaryDoc.appendChild(substituteBlock);

        // Set the tag's placeholder to the building block
        tag.setPlaceholderName("My placeholder");

        Assert.assertEquals("Custom placeholder text.", msString.trim(tag.getText()));
        Assert.assertTrue(tag.isShowingPlaceholderText());

        // Edit the text of the structured document tag and disable showing of placeholder text
        Run run = (Run)tag.getChild(NodeType.RUN, 0, true);
        run.setText("New text.");
        tag.isShowingPlaceholderText(false);

        Assert.assertEquals("New text.", msString.trim(tag.getText()));

        tag.clear();

        // Clearing a PlainText tag reverts these changes
        Assert.assertTrue(tag.isShowingPlaceholderText());
        Assert.assertEquals("Custom placeholder text.", msString.trim(tag.getText()));
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
        //ExSummary:Shows how to insert a StructuredDocumentTag as a building block and set its category and gallery.
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
        //ExSummary:Shows how structured document tags can be updated while saving to .pdf.
        Document doc = new Document();

        // Insert two StructuredDocumentTags; a date and a drop down list 
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.DATE, MarkupLevel.BLOCK);
        tag.setFullDateInternal(DateTime.getNow());

        doc.getFirstSection().getBody().appendChild(tag);

        tag = new StructuredDocumentTag(doc, SdtType.DROP_DOWN_LIST, MarkupLevel.BLOCK);
        tag.getListItems().add(new SdtListItem("Value 1"));
        tag.getListItems().add(new SdtListItem("Value 2"));
        tag.getListItems().add(new SdtListItem("Value 3"));
        tag.getListItems().setSelectedValue(tag.getListItems().get(1));

        doc.getFirstSection().getBody().appendChild(tag);

        // We've selected default values for both tags
        // We can save those values in the document without immediately updating the tags, leaving them in their default state
        // by using a SaveOptions object with this flag set
        PdfSaveOptions options = new PdfSaveOptions();
        options.setUpdateSdtContent(updateSdtContent);

        doc.save(getArtifactsDir() + "StructuredDocumentTag.UpdateSdtContent.pdf", options);
        //ExEnd

        Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(getArtifactsDir() + "StructuredDocumentTag.UpdateSdtContent.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();
        textAbsorber.Visit(pdfDoc);

        Assert.AreEqual(updateSdtContent ? "Value 2" : "Click here to enter a date.\r\nChoose an item.",
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
        //ExSummary:Shows how to fill the table with data contained in the XML part.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
 
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add("Books",
            "<books>" +
            "<book><title>Everyday Italian</title>" +
            "<author>Giada De Laurentiis</author></book>" +
            "<book><title>Harry Potter</title>" +
            "<author>J. K. Rowling</author></book>" +
            "<book><title>Learning XML</title>" +
            "<author>Erik T. Ray</author></book>" +
            "</books>");
 
        // Create headers for data from xml content
        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Title");
        builder.insertCell();
        builder.write("Author");
        builder.endRow();
        builder.endTable();
 
        // Create table with RepeatingSection inside
        StructuredDocumentTag repeatingSectionSdt =
            new StructuredDocumentTag(doc, SdtType.REPEATING_SECTION, MarkupLevel.ROW);
        repeatingSectionSdt.getXmlMapping().setMapping(xmlPart, "/books[1]/book", "");
        table.appendChild(repeatingSectionSdt);
 
        // Add RepeatingSectionItem inside RepeatingSection and mark it as a row
        StructuredDocumentTag repeatingSectionItemSdt =
            new StructuredDocumentTag(doc, SdtType.REPEATING_SECTION_ITEM, MarkupLevel.ROW);
        repeatingSectionSdt.appendChild(repeatingSectionItemSdt);
 
        Row row = new Row(doc);
        repeatingSectionItemSdt.appendChild(row);
 
        // Map xml data with created table cells for book title and author
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
                        "Harry Potter\u0007J. K. Rowling\u0007\u0007" +
                        "Learning XML\u0007Erik T. Ray\u0007\u0007", msString.trim(doc.getChild(NodeType.TABLE, 0, true).getText()));
    }

    @Test
    public void customXmlPart() throws Exception
    {
        // Obtain an XML in the form of a string
        String xmlString = "<?xml version=\"1.0\"?>" +
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

        // Insert the full XML document as a custom document part
        // The mapping for this part will be seen in the "XML Mapping Pane" in the "Developer" tab, if it is enabled
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add(Guid.newGuid().toString("B"), xmlString);

        // None of the XML is in the document body at this point
        // Create a StructuredDocumentTag, which will refer to a single element from the XML with an XPath
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.BLOCK);
        sdt.getXmlMapping().setMapping(xmlPart, "Company//Employee[@id='2']/FirstName", "");

        // Add the StructuredDocumentTag to the document to display the element in the text 
        doc.getFirstSection().getBody().appendChild(sdt);
    }
}
