// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.StructuredDocumentTag;
import com.aspose.ms.System.msConsole;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.words.SdtType;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Style;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.MarkupLevel;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.SaveFormat;
import com.aspose.words.Node;
import com.aspose.words.CustomXmlPart;
import com.aspose.ms.System.Guid;


/// <summary>
/// Tests that verify work with structured document tags in the document 
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
        Document doc = new Document(getMyDir() + "TestRepeatingSection.docx");

        NodeCollection sdTags = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        for (StructuredDocumentTag sdTag : sdTags.<StructuredDocumentTag>OfType() !!Autoporter error: Undefined expression type )
        {
            msConsole.writeLine("Type of this SDT is: {0}", sdTag.getSdtType());
        }

        //ExEnd
        StructuredDocumentTag sdTagRepeatingSection = (StructuredDocumentTag) sdTags.get(0);
        msAssert.areEqual(SdtType.REPEATING_SECTION, sdTagRepeatingSection.getSdtType());

        StructuredDocumentTag sdTagRichText = (StructuredDocumentTag) sdTags.get(1);
        msAssert.areEqual(SdtType.RICH_TEXT, sdTagRichText.getSdtType());
    }

    @Test
    public void setSpecificStyleToSdt() throws Exception
    {
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

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        NodeCollection tags = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        for (Node node : (Iterable<Node>) tags)
        {
            StructuredDocumentTag sdt = (StructuredDocumentTag) node;
            // If style was not defined before, style should be "Default Paragraph Font"
            msAssert.areEqual(StyleIdentifier.QUOTE, sdt.getStyle().getStyleIdentifier());
            msAssert.areEqual("Quote", sdt.getStyleName());
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
        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        NodeCollection sdts = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        StructuredDocumentTag sdt = (StructuredDocumentTag) sdts.get(0);
        msAssert.areEqual(true, sdt.getChecked());
        Assert.That(sdt.getXmlMapping().getStoreItemId(), Is.Empty); //Assert that this sdt has no StoreItemId
    }

    @Test (groups = "SkipTearDown")
    public void creatingCustomXml() throws Exception
    {
        //ExStart
        //ExFor:CustomXmlPart
        //ExFor:CustomXmlPartCollection.Add(String, String)
        //ExFor:Document.CustomXmlParts
        //ExFor:XmlMapping.SetMapping(CustomXmlPart, String, String)
        //ExSummary:Shows how to create structured document tag with a custom XML data.
        Document doc = new Document();
        // Add test XML data part to the collection.
        CustomXmlPart xmlPart =
            doc.getCustomXmlParts().add(Guid.newGuid().toString("B"), "<root><text>Hello, World!</text></root>");

        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.BLOCK);
        sdt.getXmlMapping().setMapping(xmlPart, "/root[1]/text[1]", "");

        doc.getFirstSection().getBody().appendChild(sdt);

        doc.save(getArtifactsDir() + "SDT.CustomXml.docx");
        //ExEnd

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "SDT.CustomXml.docx", getGoldsDir() + "SDT.CustomXml Gold.docx"));
    }

    @Test
    public void customXmlPartStoreItemIdReadOnly() throws Exception
    {
        //ExStart
        //ExFor:XmlMapping.StoreItemId
        //ExSummary:Shows how to get special id of your xml part.
        Document doc = new Document(getArtifactsDir() + "SDT.CustomXml.docx");

        StructuredDocumentTag sdt = (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        msConsole.writeLine("The Id of your custom xml part is: " + sdt.getXmlMapping().getStoreItemId());
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
        
        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        StructuredDocumentTag sdt = (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        msConsole.writeLine("The Id of your custom xml part is: " + sdt.getXmlMapping().getStoreItemId());
    }

    @Test
    public void clearTextFromStructuredDocumentTags() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.Clear
        //ExSummary:Shows how to delete content of StructuredDocumentTag elements.
        Document doc = new Document(getMyDir() + "TestRepeatingSection.docx");

        NodeCollection sdts = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);
        Assert.assertNotNull(sdts);

        for (StructuredDocumentTag sdt : sdts.<StructuredDocumentTag>OfType() !!Autoporter error: Undefined expression type )
        {
            sdt.clear();
        }

        //ExEnd
        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        sdts = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        msAssert.areEqual(
            "Enter any content that you want to repeat, including other content controls. You can also insert this control around table rows in order to repeat parts of a table.\r",
            sdts.get(0).getText());
        msAssert.areEqual("Click here to enter text.\f", sdts.get(2).getText());
    }

    @Test
    public void accessToBuildingBlockPropertiesFromDocPartObjSdt() throws Exception
    {
        Document doc = new Document(getMyDir() + "StructuredDocumentTag.BuildingBlocks.docx");

        StructuredDocumentTag docPartObjSdt =
            (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);

        msAssert.areEqual(SdtType.DOC_PART_OBJ, docPartObjSdt.getSdtType());
        msAssert.areEqual("Table of Contents", docPartObjSdt.getBuildingBlockGallery());
    }

    @Test
    public void accessToBuildingBlockPropertiesFromPlainTextSdt() throws Exception
    {
        Document doc = new Document(getMyDir() + "StructuredDocumentTag.BuildingBlocks.docx");

        StructuredDocumentTag plainTextSdt =
            (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 1, true);

        msAssert.areEqual(SdtType.PLAIN_TEXT, plainTextSdt.getSdtType());
        Assert.That(() => plainTextSdt.getBuildingBlockGallery(), Throws.<IllegalStateException>TypeOf(),
            "BuildingBlockType is only accessible for BuildingBlockGallery SDT type.");
    }

    @Test
    public void accessToBuildingBlockPropertiesFromBuildingBlockGallerySdtType() throws Exception
    {
        Document doc = new Document();

        StructuredDocumentTag buildingBlockSdt =
            new StructuredDocumentTag(doc, SdtType.BUILDING_BLOCK_GALLERY, MarkupLevel.BLOCK);
            {
                buildingBlockSdt.setBuildingBlockCategory("Built-in");
                buildingBlockSdt.setBuildingBlockGallery("Table of Contents");
            }

        doc.getFirstSection().getBody().appendChild(buildingBlockSdt);

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        buildingBlockSdt =
            (StructuredDocumentTag) doc.getFirstSection().getBody().getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);

        msAssert.areEqual(SdtType.BUILDING_BLOCK_GALLERY, buildingBlockSdt.getSdtType());
        msAssert.areEqual("Table of Contents", buildingBlockSdt.getBuildingBlockGallery());
        msAssert.areEqual("Built-in", buildingBlockSdt.getBuildingBlockCategory());
    }
}
