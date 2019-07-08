package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.ByteArrayOutputStream;
import java.text.MessageFormat;
import java.util.UUID;

/**
 * Tests that verify work with structured document tags in the document
 */
@Test
public class ExStructuredDocumentTag extends ApiExampleBase {
    @Test
    public void repeatingSection() throws Exception {
        //ExStart
        //ExFor:StructuredDocumentTag.SdtType
        //ExSummary:Shows how to get type of structured document tag.
        Document doc = new Document(getMyDir() + "TestRepeatingSection.docx");

        NodeCollection sdTags = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        for (StructuredDocumentTag sdTag : (Iterable<StructuredDocumentTag>) sdTags) {
            System.out.println(MessageFormat.format("Type of this SDT is: {0}", sdTag.getSdtType()));
        }
        //ExEnd

        StructuredDocumentTag sdTagRepeatingSection = (StructuredDocumentTag) sdTags.get(0);
        Assert.assertEquals(sdTagRepeatingSection.getSdtType(), SdtType.REPEATING_SECTION);

        StructuredDocumentTag sdTagRichText = (StructuredDocumentTag) sdTags.get(1);
        Assert.assertEquals(sdTagRichText.getSdtType(), SdtType.RICH_TEXT);
    }

    @Test
    public void setSpecificStyleToSdt() throws Exception {
        //ExStart
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
        sdtRichText.setStyleName("Quote"); // Second method to apply specific style to an SDT control

        // Insert content controls into the document
        builder.insertNode(sdtPlainText);
        builder.insertNode(sdtRichText);

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

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
    }

    @Test
    public void creatingCustomXml() throws Exception {
        //ExStart
        //ExFor:CustomXmlPart
        //ExFor:CustomXmlPartCollection.Add(String, String)
        //ExFor:Document.CustomXmlParts
        //ExFor:XmlMapping.SetMapping(CustomXmlPart, String, String)
        //ExSummary:Shows how to create structured document tag with a custom XML data.
        Document doc = new Document();
        // Add test XML data part to the collection.
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add(UUID.randomUUID().toString(), "<root><text>Hello, World!</text></root>");

        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.BLOCK);
        sdt.getXmlMapping().setMapping(xmlPart, "/root[1]/text[1]", "");

        doc.getFirstSection().getBody().appendChild(sdt);

        doc.save(getArtifactsDir() + "SDT.CustomXml.docx");
        //ExEnd

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "SDT.CustomXml.docx", getGoldsDir() + "SDT.CustomXml Gold.docx"));
    }

    @Test
    public void customXmlPartStoreItemIdReadOnly() throws Exception {
        //ExStart
        //ExFor:XmlMapping.StoreItemId
        //ExSummary:Shows how to get special id of your xml part.
        Document doc = new Document(getArtifactsDir() + "SDT.CustomXml.docx");

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
        Document doc = new Document(getMyDir() + "TestRepeatingSection.docx");

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
    public void accessToBuildingBlockPropertiesFromDocPartObjSdt() throws Exception {
        Document doc = new Document(getMyDir() + "StructuredDocumentTag.BuildingBlocks.docx");

        StructuredDocumentTag docPartObjSdt =
                (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);

        Assert.assertEquals(docPartObjSdt.getSdtType(), SdtType.DOC_PART_OBJ);
        Assert.assertEquals(docPartObjSdt.getBuildingBlockGallery(), "Table of Contents");
    }

    @Test(expectedExceptions = IllegalStateException.class)
    public void accessToBuildingBlockPropertiesFromPlainTextSdt() throws Exception {
        Document doc = new Document(getMyDir() + "StructuredDocumentTag.BuildingBlocks.docx");

        StructuredDocumentTag plainTextSdt =
                (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 1, true);
        Assert.assertEquals(plainTextSdt.getSdtType(), SdtType.PLAIN_TEXT);

        plainTextSdt.getBuildingBlockGallery();
    }

    @Test
    public void accessToBuildingBlockPropertiesFromBuildingBlockGallerySdtType() throws Exception {
        Document doc = new Document();

        StructuredDocumentTag buildingBlockSdt =
                new StructuredDocumentTag(doc, SdtType.BUILDING_BLOCK_GALLERY, MarkupLevel.BLOCK);
        buildingBlockSdt.setBuildingBlockCategory("Built-in");
        buildingBlockSdt.setBuildingBlockGallery("Table of Contents");

        doc.getFirstSection().getBody().appendChild(buildingBlockSdt);

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        buildingBlockSdt =
                (StructuredDocumentTag) doc.getFirstSection().getBody().getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);

        Assert.assertEquals(SdtType.BUILDING_BLOCK_GALLERY, buildingBlockSdt.getSdtType());
        Assert.assertEquals("Table of Contents", buildingBlockSdt.getBuildingBlockGallery());
        Assert.assertEquals("Built-in", buildingBlockSdt.getBuildingBlockCategory());
    }
}
