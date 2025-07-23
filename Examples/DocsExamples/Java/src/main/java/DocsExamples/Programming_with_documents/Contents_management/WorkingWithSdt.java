package DocsExamples.Programming_with_documents.Contents_management;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Shape;
import com.aspose.words.*;
import org.testng.annotations.Test;

import java.awt.*;
import java.nio.charset.StandardCharsets;
import java.util.UUID;

@Test
public class WorkingWithSdt extends DocsExamplesBase
{
    @Test
    public void sdtCheckBox() throws Exception
    {
        //ExStart:SdtCheckBox
        //GistId:089defec1b191de967e6099effeabda7
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        StructuredDocumentTag sdtCheckBox = new StructuredDocumentTag(doc, SdtType.CHECKBOX, MarkupLevel.INLINE);
        builder.insertNode(sdtCheckBox);
        
        doc.save(getArtifactsDir() + "WorkingWithSdt.SdtCheckBox.docx", SaveFormat.DOCX);
        //ExEnd:SdtCheckBox
    }

    @Test
    public void currentStateOfCheckBox() throws Exception
    {
        //ExStart:CurrentStateOfCheckBox
        //GistId:089defec1b191de967e6099effeabda7
        Document doc = new Document(getMyDir() + "Structured document tags.docx");
        
        // Get the first content control from the document.
        StructuredDocumentTag sdtCheckBox =
            (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);

        if (sdtCheckBox.getSdtType() == SdtType.CHECKBOX)
            sdtCheckBox.setChecked(true);

        doc.save(getArtifactsDir() + "WorkingWithSdt.CurrentStateOfCheckBox.docx");
        //ExEnd:CurrentStateOfCheckBox
    }

    @Test
    public void modifySdt() throws Exception
    {
        //ExStart:ModifySdt
        //GistId:089defec1b191de967e6099effeabda7
        Document doc = new Document(getMyDir() + "Structured document tags.docx");

        for (StructuredDocumentTag sdt : (Iterable<StructuredDocumentTag>) doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true))
        {
            switch (sdt.getSdtType())
            {
                case SdtType.PLAIN_TEXT:
                {
                    sdt.removeAllChildren();
                    Paragraph para = (Paragraph) sdt.appendChild(new Paragraph(doc));
                    Run run = new Run(doc, "new text goes here");
                    para.appendChild(run);
                    break;
                }
                case SdtType.DROP_DOWN_LIST:
                {
                    SdtListItem secondItem = sdt.getListItems().get(2);
                    sdt.getListItems().setSelectedValue(secondItem);
                    break;
                }
                case SdtType.PICTURE:
                {
                    Shape shape = (Shape) sdt.getChild(NodeType.SHAPE, 0, true);
                    if (shape.hasImage())
                    {
                        shape.getImageData().setImage(getImagesDir() + "Watermark.png");
                    }

                    break;
                }
            }
        }
        
        doc.save(getArtifactsDir() + "WorkingWithSdt.ModifySdt.docx");
        //ExEnd:ModifySdt
    }

    @Test
    public void sdtComboBox() throws Exception
    {
        //ExStart:SdtComboBox
        //GistId:089defec1b191de967e6099effeabda7
        Document doc = new Document();

        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.COMBO_BOX, MarkupLevel.BLOCK);
        sdt.getListItems().add(new SdtListItem("Choose an item", "-1"));
        sdt.getListItems().add(new SdtListItem("Item 1", "1"));
        sdt.getListItems().add(new SdtListItem("Item 2", "2"));
        doc.getFirstSection().getBody().appendChild(sdt);

        doc.save(getArtifactsDir() + "WorkingWithSdt.SdtComboBox.docx");
        //ExEnd:SdtComboBox
    }

    @Test
    public void sdtRichTextBox() throws Exception
    {
        //ExStart:SdtRichTextBox
        //GistId:089defec1b191de967e6099effeabda7
        Document doc = new Document();

        StructuredDocumentTag sdtRichText = new StructuredDocumentTag(doc, SdtType.RICH_TEXT, MarkupLevel.BLOCK);

        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc);
        run.setText("Hello World");
        run.getFont().setColor(Color.GREEN);
        para.getRuns().add(run);
        sdtRichText.getChildNodes(NodeType.ANY, false).add(para);
        doc.getFirstSection().getBody().appendChild(sdtRichText);

        doc.save(getArtifactsDir() + "WorkingWithSdt.SdtRichTextBox.docx");
        //ExEnd:SdtRichTextBox
    }

    @Test
    public void sdtColor() throws Exception
    {
        //ExStart:SdtColor
        //GistId:089defec1b191de967e6099effeabda7
        Document doc = new Document(getMyDir() + "Structured document tags.docx");

        StructuredDocumentTag sdt = (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        sdt.setColor(Color.RED);

        doc.save(getArtifactsDir() + "WorkingWithSdt.SdtColor.docx");
        //ExEnd:SdtColor
    }

    @Test
    public void clearSdt() throws Exception
    {
        //ExStart:ClearSdt
        //GistId:089defec1b191de967e6099effeabda7
        Document doc = new Document(getMyDir() + "Structured document tags.docx");

        StructuredDocumentTag sdt = (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        sdt.clear();

        doc.save(getArtifactsDir() + "WorkingWithSdt.ClearSdt.doc");
        //ExEnd:ClearSdt
    }

    @Test
    public void bindSdtToCustomXmlPart() throws Exception
    {
        //ExStart:BindSdtToCustomXmlPart
        //GistId:089defec1b191de967e6099effeabda7
        Document doc = new Document();
        CustomXmlPart xmlPart =
            doc.getCustomXmlParts().add(UUID.randomUUID().toString(), "<root><text>Hello, World!</text></root>");

        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.BLOCK);
        doc.getFirstSection().getBody().appendChild(sdt);

        sdt.getXmlMapping().setMapping(xmlPart, "/root[1]/text[1]", "");

        doc.save(getArtifactsDir() + "WorkingWithSdt.BindSdtToCustomXmlPart.doc");
        //ExEnd:BindSdtToCustomXmlPart
    }

    @Test
    public void sdtStyle() throws Exception
    {
        //ExStart:SdtStyle
        //GistId:089defec1b191de967e6099effeabda7
        Document doc = new Document(getMyDir() + "Structured document tags.docx");

        StructuredDocumentTag sdt = (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        Style style = doc.getStyles().getByStyleIdentifier(StyleIdentifier.QUOTE);
        sdt.setStyle(style);

        doc.save(getArtifactsDir() + "WorkingWithSdt.SdtStyle.docx");
        //ExEnd:SdtStyle
    }

    @Test
    public void repeatingSectionMappedToCustomXmlPart() throws Exception
    {
        //ExStart:RepeatingSectionMappedToCustomXmlPart
        //GistId:089defec1b191de967e6099effeabda7
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        CustomXmlPart xmlPart = doc.getCustomXmlParts().add("Books",
            "<books><book><title>Everyday Italian</title><author>Giada De Laurentiis</author></book>" +
            "<book><title>Harry Potter</title><author>J K. Rowling</author></book>" +
            "<book><title>Learning XML</title><author>Erik T. Ray</author></book></books>");

        Table table = builder.startTable();

        builder.insertCell();
        builder.write("Title");

        builder.insertCell();
        builder.write("Author");

        builder.endRow();
        builder.endTable();

        StructuredDocumentTag repeatingSectionSdt =
            new StructuredDocumentTag(doc, SdtType.REPEATING_SECTION, MarkupLevel.ROW);
        repeatingSectionSdt.getXmlMapping().setMapping(xmlPart, "/books[1]/book", "");
        table.appendChild(repeatingSectionSdt);

        StructuredDocumentTag repeatingSectionItemSdt = 
            new StructuredDocumentTag(doc, SdtType.REPEATING_SECTION_ITEM, MarkupLevel.ROW);
        repeatingSectionSdt.appendChild(repeatingSectionItemSdt);

        Row row = new Row(doc);
        repeatingSectionItemSdt.appendChild(row);

        StructuredDocumentTag titleSdt =
            new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.CELL);
        titleSdt.getXmlMapping().setMapping(xmlPart, "/books[1]/book[1]/title[1]", "");
        row.appendChild(titleSdt);

        StructuredDocumentTag authorSdt =
            new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.CELL);
        authorSdt.getXmlMapping().setMapping(xmlPart, "/books[1]/book[1]/author[1]", "");
        row.appendChild(authorSdt);

        doc.save(getArtifactsDir() + "WorkingWithSdt.RepeatingSectionMappedToCustomXmlPart.docx");
        //ExEnd:RepeatingSectionMappedToCustomXmlPart
    }

    @Test
    public void multiSection() throws Exception
    {
        //ExStart:MultiSectionSDT
        Document doc = new Document(getMyDir() + "Multi-section structured document tags.docx");

        NodeCollection tags = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG_RANGE_START, true);

        for (StructuredDocumentTagRangeStart tag : (Iterable<StructuredDocumentTagRangeStart>) tags)
            System.out.println(tag.getTitle());
        //ExEnd:MultiSectionSDT
    }

    @Test
    public void sdtRangeStartXmlMapping() throws Exception
    {
        //ExStart:SdtRangeStartXmlMapping
        //GistId:089defec1b191de967e6099effeabda7
        Document doc = new Document(getMyDir() + "Multi-section structured document tags.docx");

        // Construct an XML part that contains data and add it to the document's CustomXmlPart collection.
        String xmlPartId = UUID.randomUUID().toString();
        String xmlPartContent = "<root><text>Text element #1</text><text>Text element #2</text></root>";
        CustomXmlPart xmlPart = doc.getCustomXmlParts().add(xmlPartId, xmlPartContent);
        System.out.println(new String(xmlPart.getData(), StandardCharsets.US_ASCII));

        // Create a StructuredDocumentTag that will display the contents of our CustomXmlPart in the document.
        StructuredDocumentTagRangeStart sdtRangeStart = (StructuredDocumentTagRangeStart)doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG_RANGE_START, 0, true);

        // If we set a mapping for our StructuredDocumentTag,
        // it will only display a part of the CustomXmlPart that the XPath points to.
        // This XPath will point to the contents second "<text>" element of the first "<root>" element of our CustomXmlPart.
        sdtRangeStart.getXmlMapping().setMapping(xmlPart, "/root[1]/text[2]", null);

        doc.save(getArtifactsDir() + "WorkingWithSdt.SdtRangeStartXmlMapping.docx");
        //ExEnd:SdtRangeStartXmlMapping
    }
}
