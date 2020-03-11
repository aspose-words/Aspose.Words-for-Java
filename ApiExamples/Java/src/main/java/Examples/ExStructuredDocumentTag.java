package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.awt.*;
import java.io.ByteArrayOutputStream;
import java.text.MessageFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.UUID;

/**
 * Tests that verify work with structured document tags in the document.
 */
@Test
public class ExStructuredDocumentTag extends ApiExampleBase {
    @Test
    public void repeatingSection() throws Exception {
        //ExStart
        //ExFor:StructuredDocumentTag.SdtType
        //ExSummary:Shows how to get type of structured document tag.
        Document doc = new Document(getMyDir() + "Structured document tags.docx");

        NodeCollection sdTags = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        for (StructuredDocumentTag sdTag : (Iterable<StructuredDocumentTag>) sdTags) {
            System.out.println(MessageFormat.format("Type of this SDT is: {0}", sdTag.getSdtType()));
        }
        //ExEnd

        StructuredDocumentTag sdTagRepeatingSection = (StructuredDocumentTag) sdTags.get(0);
        Assert.assertEquals(sdTagRepeatingSection.getSdtType(), SdtType.REPEATING_SECTION);

        StructuredDocumentTag sdTagRepeatingSectionItem = (StructuredDocumentTag) sdTags.get(1);
        Assert.assertEquals(sdTagRepeatingSectionItem.getSdtType(), SdtType.REPEATING_SECTION_ITEM);

        StructuredDocumentTag sdTagRichText = (StructuredDocumentTag) sdTags.get(2);
        Assert.assertEquals(sdTagRichText.getSdtType(), SdtType.RICH_TEXT);
    }

    @Test
    public void setSpecificStyleToSdt() throws Exception {
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

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        // We can get a collection of StructuredDocumentTags by looking for the document's child nodes of this NodeType
        Assert.assertEquals(sdtPlainText.getNodeType(), NodeType.STRUCTURED_DOCUMENT_TAG);

        NodeCollection tags = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        for (StructuredDocumentTag sdt : (Iterable<StructuredDocumentTag>) tags) {
            // If style was not defined before, style should be "Default Paragraph Font"
            Assert.assertEquals(StyleIdentifier.QUOTE, sdt.getStyle().getStyleIdentifier());
            Assert.assertEquals("Quote", sdt.getStyleName());
        }
        //ExEnd
    }

    @Test
    public void checkBox() throws Exception {
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
        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        NodeCollection sdts = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        StructuredDocumentTag sdt = (StructuredDocumentTag) sdts.get(0);
        Assert.assertEquals(sdt.getChecked(), true);
        // Assert that this sdt has no StoreItemId
        Assert.assertTrue(sdt.getXmlMapping().getStoreItemId().isEmpty());
    }

    @Test
    public void date() throws Exception {
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
        sdtDate.setDateDisplayLocale(1025);

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
        Calendar cal = Calendar.getInstance();
        cal.set(1440, 10, 20);
        sdtDate.setFullDate(cal.getTime());

        // Insert the StructuredDocumentTag into the document with a DocumentBuilder and save the document
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertNode(sdtDate);

        doc.save(getArtifactsDir() + "StructuredDocumentTag.Date.docx");
        //ExEnd
    }

    @Test
    public void plainText() throws Exception {
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
        Assert.assertTrue(tag.getId() > 0);

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
        StructuredDocumentTag tagClone = (StructuredDocumentTag) tag.deepClone(true);
        builder.insertParagraph();
        builder.insertNode(tagClone);

        // We can remove the tag while keeping its contents where they were in the Paragraph by calling RemoveSelfOnly()
        tagClone.removeSelfOnly();

        doc.save(getArtifactsDir() + "StructuredDocumentTag.PlainText.docx");
        //ExEnd
    }

    @Test
    public void isTemporary() throws Exception {
        //ExStart
        //ExFor:StructuredDocumentTag.IsTemporary
        //ExSummary:Demonstrates the effects of making a StructuredDocumentTag temporary.
        // Create a new Document
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
    }

    @Test
    public void placeholderBuildingBlock() throws Exception {
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
        Assert.assertEquals(tag.getPlaceholder(), substituteBlock);

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
    }

    @Test
    public void lock() throws Exception {
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
    }

    @Test
    public void listItemCollection() throws Exception {
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
        Assert.assertEquals(listItems.get(0).getValue(), listItems.get(0).getDisplayText());

        // Add 3 more SdtListItems with non-empty strings passed to DisplayText
        listItems.add(new SdtListItem("Item 2", "Value 2"));
        listItems.add(new SdtListItem("Item 3", "Value 3"));
        listItems.add(new SdtListItem("Item 4", "Value 4"));

        // We can obtain a count of the SdtListItems and also set the drop down list's SelectedValue attribute to
        // automatically have one of them pre-selected when we open the document in Microsoft Word
        Assert.assertEquals(listItems.getCount(), 4);
        listItems.setSelectedValue(listItems.get(3));

        Assert.assertEquals(listItems.getSelectedValue().getValue(), "Value 4");

        // We can enumerate over the collection and print each element
        Iterator<SdtListItem> enumerator = listItems.iterator();
        while (enumerator.hasNext()) {
            SdtListItem sdtListItem = enumerator.next();
            System.out.println(MessageFormat.format("List item: {0}, value: {1}", sdtListItem.getDisplayText(), sdtListItem.getValue()));
        }

        // We can also remove elements one at a time
        listItems.removeAt(3);
        Assert.assertEquals(listItems.getCount(), 3);

        // Make sure to update the SelectedValue's index if it ever ends up out of bounds before saving the document
        listItems.setSelectedValue(listItems.get(1));

        doc.save(getArtifactsDir() + "StructuredDocumentTag.ListItemCollection.docx");

        // We can clear the whole collection at once too
        listItems.clear();
        Assert.assertEquals(listItems.getCount(), 0);
        //ExEnd
    }

    @Test
    public void creatingCustomXml() throws Exception {
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
        String xmlPartId = UUID.randomUUID().toString();
        String xmlPartContent = "<root><text>Hello, World!</text></root>";
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add(xmlPartId, xmlPartContent);

        // The data we entered resides in these variables
        Assert.assertEquals(xmlPart.getData(), xmlPartContent.getBytes());
        Assert.assertEquals(xmlPart.getId(), xmlPartId);

        // XML parts can be referenced by collection index or GUID
        Assert.assertEquals(doc.getCustomXmlParts().get(0), xmlPart);
        Assert.assertEquals(doc.getCustomXmlParts().getById(xmlPartId), xmlPart);

        // Once the part is created, we can add XML schema associations like this
        xmlPart.getSchemas().add("http://www.w3.org/2001/XMLSchema");

        // We can also clone parts and insert them into the collection directly
        CustomXmlPart xmlPartClone = xmlPart.deepClone();
        xmlPartClone.setId(UUID.randomUUID().toString());
        doc.getCustomXmlParts().add(xmlPartClone);

        Assert.assertEquals(doc.getCustomXmlParts().getCount(), 2);

        // Iterate through collection with an enumerator and print the contents of each part
        Iterator<CustomXmlPart> enumerator = doc.getCustomXmlParts().iterator();
        int index = 0;
        while (enumerator.hasNext()) {
            CustomXmlPart customXmlPart = enumerator.next();
            System.out.println(MessageFormat.format("XML part index {0}, ID: {1}", index, customXmlPart.getId()));
            System.out.println(MessageFormat.format("\tContent: {0}", customXmlPart.getData()));
            index++;
        }

        // XML parts can be removed by index
        doc.getCustomXmlParts().removeAt(1);

        Assert.assertEquals(doc.getCustomXmlParts().getCount(), 1);

        // The XML part collection itself can be cloned also
        CustomXmlPartCollection customXmlParts = doc.getCustomXmlParts().deepClone();

        // And all elements can be cleared like this
        customXmlParts.clear();

        // Create a StructuredDocumentTag that will display the contents of our part,
        // insert it into the document and save the document
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.BLOCK);
        sdt.getXmlMapping().setMapping(xmlPart, "/root[1]/text[1]", "");

        doc.getFirstSection().getBody().appendChild(sdt);

        doc.save(getArtifactsDir() + "StructuredDocumentTag.CustomXml.docx");
        //ExEnd

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "StructuredDocumentTag.CustomXml.docx", getGoldsDir() + "StructuredDocumentTag.CustomXml Gold.docx"));
    }

    @Test
    public void xmlMapping() throws Exception {
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
        String xmlPartId = UUID.randomUUID().toString();
        String xmlPartContent = "<root><text>Text element #1</text><text>Text element #2</text></root>";
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add(xmlPartId, xmlPartContent);
        System.out.println(xmlPart.getData());

        // Create a StructuredDocumentTag that will display the contents of our CustomXmlPart in the document
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.BLOCK);

        // If we set a mapping for our StructuredDocumentTag,
        // it will only display a part of the CustomXmlPart that the XPath points to
        // This XPath will point to the contents second "<text>" element of the first "<root>" element of our CustomXmlPart
        sdt.getXmlMapping().setMapping(xmlPart, "/root[1]/text[2]", "xmlns:ns='http://www.w3.org/2001/XMLSchema'");

        Assert.assertTrue(sdt.getXmlMapping().isMapped());
        Assert.assertEquals(sdt.getXmlMapping().getCustomXmlPart(), xmlPart);
        Assert.assertEquals(sdt.getXmlMapping().getXPath(), "/root[1]/text[2]");
        Assert.assertEquals(sdt.getXmlMapping().getPrefixMappings(), "xmlns:ns='http://www.w3.org/2001/XMLSchema'");

        // Add the StructuredDocumentTag to the document to display the content from our CustomXmlPart
        doc.getFirstSection().getBody().appendChild(sdt);
        doc.save(getArtifactsDir() + "StructuredDocumentTag.XmlMapping.docx");
        //ExEnd
    }

    @Test
    public void customXmlSchemaCollection() throws Exception {
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

        String xmlPartId = UUID.randomUUID().toString();
        String xmlPartContent = "<root><text>Hello, World!</text></root>";
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add(xmlPartId, xmlPartContent);

        // Once the part is created, we can add XML schema associations like this,
        // and perform other collection-related operations on the list of schemas for this part
        xmlPart.getSchemas().add("http://www.w3.org/2001/XMLSchema");

        // Collections can be cloned and elements can be added
        CustomXmlSchemaCollection schemas = xmlPart.getSchemas().deepClone();
        schemas.add("http://www.w3.org/2001/XMLSchema-instance");
        schemas.add("http://schemas.microsoft.com/office/2006/metadata/contentType");

        Assert.assertEquals(schemas.getCount(), 3);
        Assert.assertEquals(schemas.indexOf(("http://schemas.microsoft.com/office/2006/metadata/contentType")), 2);

        // We can iterate over the collection with an enumerator
        Iterator<String> enumerator = schemas.iterator();
        while (enumerator.hasNext()) {
            System.out.println(enumerator.next());
        }

        // We can also remove elements by index, element, or we can clear the entire collection
        schemas.removeAt(2);
        schemas.remove("http://www.w3.org/2001/XMLSchema");
        schemas.clear();

        Assert.assertEquals(schemas.getCount(), 0);
        //ExEnd
    }

    @Test
    public void customXmlPartStoreItemIdReadOnly() throws Exception {
        //ExStart
        //ExFor:XmlMapping.StoreItemId
        //ExSummary:Shows how to get special id of your xml part.
        Document doc = new Document(getMyDir() + "Custom XML part in structured document tag.docx");

        StructuredDocumentTag sdt = (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        System.out.println("The Id of your custom xml part is: " + sdt.getXmlMapping().getStoreItemId());
        //ExEnd
    }

    @Test
    public void customXmlPartStoreItemIdReadOnlyNull() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        StructuredDocumentTag sdtCheckBox = new StructuredDocumentTag(doc, SdtType.CHECKBOX, MarkupLevel.INLINE);
        sdtCheckBox.setChecked(true);

        // Insert content control into the document
        builder.insertNode(sdtCheckBox);

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        StructuredDocumentTag sdt = (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        System.out.println("The Id of your custom xml part is: " + sdt.getXmlMapping().getStoreItemId());
    }

    @Test
    public void clearTextFromStructuredDocumentTags() throws Exception {
        //ExStart
        //ExFor:StructuredDocumentTag.Clear
        //ExSummary:Shows how to delete content of StructuredDocumentTag elements.
        Document doc = new Document(getMyDir() + "Structured document tags.docx");

        NodeCollection sdts = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);
        Assert.assertNotNull(sdts);

        for (StructuredDocumentTag sdt : (Iterable<StructuredDocumentTag>) sdts) {
            sdt.clear();
        }
        //ExEnd
        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        sdts = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        Assert.assertEquals(sdts.get(0).getText(), "Enter any content that you want to repeat, including other content controls. You can also insert this control around table rows in order to repeat parts of a table.\r");
        Assert.assertEquals(sdts.get(2).getText(), "Click here to enter text.\f");
    }

    @Test
    public void smartTagProperties() throws Exception {
        //ExStart
        //ExFor:CustomXmlProperty.Uri
        //ExFor:CustomXmlPropertyCollection
        //ExFor:CustomXmlPropertyCollection.Add(CustomXmlProperty)
        //ExFor:CustomXmlPropertyCollection.Clear
        //ExFor:CustomXmlPropertyCollection.Contains(String)
        //ExFor:CustomXmlPropertyCollection.Count
        //ExFor:CustomXmlPropertyCollection.GetEnumerator
        //ExFor:CustomXmlPropertyCollection.IndexOfKey(String)
        //ExFor:CustomXmlPropertyCollection.Item(Int32)
        //ExFor:CustomXmlPropertyCollection.Item(String)
        //ExFor:CustomXmlPropertyCollection.Remove(String)
        //ExFor:CustomXmlPropertyCollection.RemoveAt(Int32)
        //ExSummary:Shows how to work with smart tag properties to get in depth information about smart tags.
        // Open a document that contains smart tags and their collection
        Document doc = new Document(getMyDir() + "Smart tags.doc");

        // Smart tags are an older Microsoft Word feature that can automatically detect and tag
        // any parts of the text that it registers as commonly used information objects such as names, addresses, stock tickers, dates etc
        // In Word 2003, smart tags can be turned on in Tools > AutoCorrect options... > SmartTags tab
        // In our input document there are three objects that were registered as smart tags, but since they can be nested, we have 8 in this collection
        NodeCollection smartTags = doc.getChildNodes(NodeType.SMART_TAG, true);
        Assert.assertEquals(smartTags.getCount(), 8);

        // The last smart tag is of the "Date" type, which we will retrieve here
        SmartTag smartTag = (SmartTag) smartTags.get(7);

        // The Properties attribute, for some smart tags, elaborates on the text object that Word picked up as a smart tag
        // In the case of our "Date" smart tag, its properties will let us know the year, month and day within the smart tag
        CustomXmlPropertyCollection properties = smartTag.getProperties();

        // We can enumerate over the collection and print the aforementioned properties to the console
        Assert.assertEquals(properties.getCount(), 4);

        Iterator<CustomXmlProperty> enumerator = properties.iterator();
        try {
            while (enumerator.hasNext()) {
                CustomXmlProperty customXmlProperty = enumerator.next();

                System.out.println(MessageFormat.format("Property name: {0}, value: {1}", customXmlProperty.getName(), customXmlProperty.getValue()));
                Assert.assertEquals(enumerator.next().getUri(), "");
            }
        } finally {
            if (enumerator != null) enumerator.remove();
        }

        // We can also access the elements in various ways, including as a key-value pair
        Assert.assertTrue(properties.contains("Day"));
        Assert.assertEquals(properties.get("Day").getValue(), "22");
        Assert.assertEquals(properties.get(2).getValue(), "2003");
        Assert.assertEquals(properties.indexOfKey("Month"), 1);

        // We can also remove elements by name, index or clear the collection entirely
        properties.removeAt(3);
        properties.remove("Year");
        Assert.assertEquals((properties.getCount()), 2);

        properties.clear();
        Assert.assertEquals((properties.getCount()), 0);

        // We can remove the entire smart tag like this
        smartTag.remove();
        //ExEnd
    }

    //ExStart
    //ExFor:CustomXmlProperty
    //ExFor:CustomXmlProperty.#ctor(String,String,String)
    //ExFor:CustomXmlProperty.Name
    //ExFor:CustomXmlProperty.Value
    //ExFor:Markup.SmartTag
    //ExFor:Markup.SmartTag.#ctor(DocumentBase)
    //ExFor:Markup.SmartTag.Accept(DocumentVisitor)
    //ExFor:Markup.SmartTag.Element
    //ExFor:Markup.SmartTag.Properties
    //ExFor:Markup.SmartTag.Uri
    //ExSummary:Shows how to create smart tags.
    @Test //ExSkip
    public void smartTags() throws Exception {
        Document doc = new Document();
        SmartTag smartTag = new SmartTag(doc);
        smartTag.setElement("date");

        // Specify a date and set smart tag properties accordingly
        smartTag.appendChild(new Run(doc, "May 29, 2019"));

        smartTag.getProperties().add(new CustomXmlProperty("Day", "", "29"));
        smartTag.getProperties().add(new CustomXmlProperty("Month", "", "5"));
        smartTag.getProperties().add(new CustomXmlProperty("Year", "", "2019"));

        // Set the smart tag's uri to the default
        smartTag.setUri("urn:schemas-microsoft-com:office:smarttags");

        doc.getFirstSection().getBody().getFirstParagraph().appendChild(smartTag);
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(new Run(doc, " is a date. "));

        // Create and add one more smart tag, this time for a financial symbol
        smartTag = new SmartTag(doc);
        smartTag.setElement("stockticker");
        smartTag.setUri("urn:schemas-microsoft-com:office:smarttags");

        smartTag.appendChild(new Run(doc, "MSFT"));

        doc.getFirstSection().getBody().getFirstParagraph().appendChild(smartTag);
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(new Run(doc, " is a stock ticker."));

        // Print all the smart tags in our document with a document visitor
        doc.accept(new SmartTagVisitor());

        doc.save(getArtifactsDir() + "StructuredDocumentTag.SmartTags.docx");
    }

    /// <summary>
    /// DocumentVisitor implementation that prints smart tags and their contents.
    /// </summary>
    private static class SmartTagVisitor extends DocumentVisitor {
        /// <summary>
        /// Called when a SmartTag node is encountered in the document.
        /// </summary>
        public int visitSmartTagStart(SmartTag smartTag) {
            System.out.println(MessageFormat.format("Smart tag type: {0}", smartTag.getElement()));
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a SmartTag node is ended.
        /// </summary>
        public int visitSmartTagEnd(SmartTag smartTag) throws Exception {
            System.out.println(MessageFormat.format("\tContents: \"{0}\"", smartTag.toString(SaveFormat.TEXT)));

            if (smartTag.getProperties().getCount() == 0) {
                System.out.println("\tContains no properties");

            } else {
                System.out.println("\tProperties: ");
                String[] properties = new String[smartTag.getProperties().getCount()];
                int index = 0;

                for (CustomXmlProperty cxp : smartTag.getProperties()) {
                    properties[index++] = MessageFormat.format("\"{0}\" = \"{1}\"", cxp.getName(), cxp.getValue());
                }

                System.out.println(String.join(", ", properties));
            }

            return VisitorAction.CONTINUE;
        }
    }
    //ExEnd

    @Test
    public void accessToBuildingBlockPropertiesFromDocPartObjSdt() throws Exception {
        Document doc = new Document(getMyDir() + "Structured document tags with building blocks.docx");

        StructuredDocumentTag docPartObjSdt =
                (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);

        Assert.assertEquals(docPartObjSdt.getSdtType(), SdtType.DOC_PART_OBJ);
        Assert.assertEquals(docPartObjSdt.getBuildingBlockGallery(), "Table of Contents");
    }

    @Test(expectedExceptions = IllegalStateException.class)
    public void accessToBuildingBlockPropertiesFromPlainTextSdt() throws Exception {
        Document doc = new Document(getMyDir() + "Structured document tags with building blocks.docx");

        StructuredDocumentTag plainTextSdt =
                (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 1, true);
        Assert.assertEquals(plainTextSdt.getSdtType(), SdtType.PLAIN_TEXT);

        plainTextSdt.getBuildingBlockGallery();
    }

    @Test
    public void buildingBlockCategories() throws Exception {
        //ExStart
        //ExFor:StructuredDocumentTag.BuildingBlockCategory
        //ExFor:StructuredDocumentTag.BuildingBlockGallery
        //ExSummary:Shows how to insert a StructuredDocumentTag as a building block and set its category and gallery.
        Document doc = new Document();

        StructuredDocumentTag buildingBlockSdt =
                new StructuredDocumentTag(doc, SdtType.BUILDING_BLOCK_GALLERY, MarkupLevel.BLOCK);
        buildingBlockSdt.setBuildingBlockCategory("Built-in");
        buildingBlockSdt.setBuildingBlockGallery("Table of Contents");

        doc.getFirstSection().getBody().appendChild(buildingBlockSdt);

        doc.save(getArtifactsDir() + "StructuredDocumentTag.BuildingBlockCategories.docx");
        //ExEnd

        buildingBlockSdt =
                (StructuredDocumentTag) doc.getFirstSection().getBody().getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);

        Assert.assertEquals(SdtType.BUILDING_BLOCK_GALLERY, buildingBlockSdt.getSdtType());
        Assert.assertEquals("Table of Contents", buildingBlockSdt.getBuildingBlockGallery());
        Assert.assertEquals("Built-in", buildingBlockSdt.getBuildingBlockCategory());
    }

    @Test
    public void updateSdtContent() throws Exception {
        //ExStart
        //ExFor:SaveOptions.UpdateSdtContent
        //ExSummary:Shows how structured document tags can be updated while saving to .pdf.
        Document doc = new Document();

        // Insert two StructuredDocumentTags; a date and a drop down list
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.DATE, MarkupLevel.BLOCK);
        tag.setFullDate(new Date());

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
        options.setUpdateSdtContent(false);

        doc.save(getArtifactsDir() + "StructuredDocumentTag.UpdateSdtContent.pdf", options);
        //ExEnd
    }

    @Test
    public void fillTableUsingRepeatingSectionItem() throws Exception {
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
                        "<author>J K. Rowling</author></book>" +
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
    }

    @Test
    public void customXmlPart() throws Exception {
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

        // Create a blank document
        Document doc = new Document();

        // Insert the full XML document as a custom document part
        // The mapping for this part will be seen in the "XML Mapping Pane" in the "Developer" tab, if it is enabled
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add(UUID.randomUUID().toString(), xmlString);

        // None of the XML is in the document body at this point
        // Create a StructuredDocumentTag, which will refer to a single element from the XML with an XPath
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.BLOCK);
        sdt.getXmlMapping().setMapping(xmlPart, "Company//Employee[@id='2']/FirstName", "");

        // Add the StructuredDocumentTag to the document to display the element in the text
        doc.getFirstSection().getBody().appendChild(sdt);
    }
}
