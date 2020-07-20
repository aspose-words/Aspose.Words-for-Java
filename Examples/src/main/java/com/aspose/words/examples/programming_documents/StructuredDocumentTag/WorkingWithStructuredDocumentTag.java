package com.aspose.words.examples.programming_documents.StructuredDocumentTag;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;
import com.aspose.words.List;
import com.aspose.words.StructuredDocumentTagRangeStart;

import java.awt.*;

public class WorkingWithStructuredDocumentTag {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(WorkingWithStructuredDocumentTag.class);

        setContentControlColor(dataDir);
        setContentControlStyle(dataDir);
        CreatingTableRepeatingSectionMappedToCustomXmlPart(dataDir);
        MultiSectionSDT(dataDir);
    }

    public static void setContentControlColor(String dataDir) throws Exception {
        // ExStart:SetContentControlColor
        // The path to the documents directory.
        Document doc = new Document(dataDir + "input.docx");
        StructuredDocumentTag sdt = (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        sdt.setColor(Color.RED);

        dataDir = dataDir + "SetContentControlColor_out.docx";

        // Save the document to disk.
        doc.save(dataDir);
        // ExEnd:SetContentControlColor
        System.out.println("\nSet the color of content control successfully.");
    }

    public static void setContentControlStyle(String dataDir) throws Exception {
        // ExStart:setContentControlStyle
        Document doc = new Document(dataDir + "input.docx");
        StructuredDocumentTag sdt = (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        Style style = doc.getStyles().getByStyleIdentifier(StyleIdentifier.QUOTE);
        sdt.setStyle(style);
        dataDir = dataDir + "SetContentControlStyle_out.docx";
        doc.save(dataDir);
        // ExEnd:setContentControlStyle
        System.out.println("\nSet the style of content control successfully.");
    }
    
    public static void CreatingTableRepeatingSectionMappedToCustomXmlPart(String dataDir) throws Exception
    {
        // ExStart:CreatingTableRepeatingSectionMappedToCustomXmlPart
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

        doc.save(dataDir + "Document.docx");
        // ExEnd:CreatingTableRepeatingSectionMappedToCustomXmlPart
        System.out.println("\nCreation of a Table Repeating Section Mapped To a Custom Xml Part is successfull.");
    }
    
    public static void MultiSectionSDT(String dataDir) throws Exception {
    	// ExStart:MultiSectionSDT
    	Document doc = new Document(dataDir + "input.docx");
    	NodeCollection<StructuredDocumentTagRangeStart> tags = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG_RANGE_START, true);

    	for (StructuredDocumentTagRangeStart tag : tags)
    	    System.out.println(tag.getTitle());
    	// ExEnd:MultiSectionSDT
    }
}
