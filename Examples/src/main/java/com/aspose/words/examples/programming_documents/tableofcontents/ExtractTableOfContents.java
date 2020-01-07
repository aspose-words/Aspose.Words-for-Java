package com.aspose.words.examples.programming_documents.tableofcontents;

import java.util.ArrayList;

import com.aspose.words.Bookmark;
import com.aspose.words.Document;
import com.aspose.words.Field;
import com.aspose.words.FieldHyperlink;
import com.aspose.words.FieldStart;
import com.aspose.words.FieldType;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class ExtractTableOfContents {

	public static void main(String[] args) throws Exception {
		//ExStart:ExtractTableOfContents
		// The path to the documents directory.
        String dataDir = Utils.getSharedDataDir(ExtractTableOfContents.class) + "TableOfContents/";

        String fileName = "TOC.doc";
        Document doc = new Document(dataDir + fileName);
        
        for (Field field : (Iterable<Field>)doc.getRange().getFields())
        {
            if (field.getType() == FieldType.FIELD_HYPERLINK)
            {
                FieldHyperlink hyperlink = (FieldHyperlink)field;
                if (hyperlink.getSubAddress() != null && hyperlink.getSubAddress().startsWith("_Toc"))
                {
                    Paragraph tocItem = (Paragraph)field.getStart().getAncestor(NodeType.PARAGRAPH);
                    System.out.println(tocItem.toString(SaveFormat.TEXT).trim());
                    System.out.println("------------------");
                    if (tocItem != null)
                    {
                        Bookmark bm = doc.getRange().getBookmarks().get(hyperlink.getSubAddress());
                        // Get the location this TOC Item is pointing to
                        Paragraph pointer = (Paragraph)bm.getBookmarkStart().getAncestor(NodeType.PARAGRAPH);
                        System.out.println(pointer.toString(SaveFormat.TEXT));
                    }
                } // End If
            }// End If
        }// End Foreach
      //ExEnd:ExtractTableOfContents
	}

}
