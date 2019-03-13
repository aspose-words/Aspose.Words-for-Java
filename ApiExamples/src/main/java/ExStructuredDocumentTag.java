//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.StructuredDocumentTag;
import org.testng.Assert;
import com.aspose.words.SdtType;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.MarkupLevel;
import com.aspose.words.SaveFormat;
import com.aspose.words.CustomXmlPart;

import java.io.ByteArrayOutputStream;
import java.text.MessageFormat;
import java.util.UUID;

/**
 *  Tests that verify work with structured document tags in the document 
 */
@Test
public class ExStructuredDocumentTag extends ApiExampleBase
{
    @Test
    public void repeatingSection() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.SdtType
        //ExSummary:Shows how to get type of structured document tag.
        Document doc = new Document(getMyDir() + "TestRepeatingSection.docx");

        NodeCollection sdTags = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        for (StructuredDocumentTag sdTag : (Iterable<StructuredDocumentTag>) sdTags)
        {
            System.out.println(MessageFormat.format("Type of this SDT is: {0}", sdTag.getSdtType()));
        }
        //ExEnd
        StructuredDocumentTag sdTagRepeatingSection = (StructuredDocumentTag)sdTags.get(0);
        Assert.assertEquals(sdTagRepeatingSection.getSdtType(), SdtType.REPEATING_SECTION);

        StructuredDocumentTag sdTagRichText = (StructuredDocumentTag)sdTags.get(1);
        Assert.assertEquals(sdTagRichText.getSdtType(), SdtType.RICH_TEXT);
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
        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        NodeCollection sdts = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        StructuredDocumentTag sdt = (StructuredDocumentTag)sdts.get(0);
        Assert.assertEquals(sdt.getChecked(), true);
    }

    @Test
    public void creatingCustomXml() throws Exception
    {
        //ExStart
        //ExFor:CustomXmlPart
        //ExFor:CustomXmlPartCollection.Add(String, String)
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
    public void clearTextFromStructuredDocumentTags() throws Exception
    {
        //ExStart
        //ExFor:StructuredDocumentTag.Clear
        //ExSummary:Shows how to delete content of StructuredDocumentTag elements.
        Document doc = new Document(getMyDir() + "TestRepeatingSection.docx");

        NodeCollection sdts = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);
        Assert.assertNotNull(sdts);

        for (StructuredDocumentTag sdt : (Iterable<StructuredDocumentTag>) sdts)
        {
            sdt.clear();
        }
        //ExEnd
        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        sdts = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        Assert.assertEquals(sdts.get(0).getText(), "Enter any content that you want to repeat, including other content controls. You can also insert this control around table rows in order to repeat parts of a table.\r");
        Assert.assertEquals(sdts.get(2).getText(), "Click here to enter text.\f");
    }
}
