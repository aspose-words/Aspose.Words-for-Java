/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.util.HashMap;


public class ListUseDestinationStyles
{
    private static String gDataDir;

    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        gDataDir = Utils.getDataDir(ListUseDestinationStyles.class);

        Document dstDoc = new Document(gDataDir + "TestFile.DestinationList.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.SourceList.doc");

        // Set the source document to continue straight after the end of the destination document.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        // Keep track of the lists that are created.
        HashMap newLists = new HashMap();

        // Iterate through all paragraphs in the document.
        for (Paragraph para : (Iterable<Paragraph>) srcDoc.getChildNodes(NodeType.PARAGRAPH, true))
        {
            if (para.isListItem())
            {
                int listId = para.getListFormat().getList().getListId();

                // Check if the destination document contains a list with this ID already. If it does then this may
                // cause the two lists to run together. Create a copy of the list in the source document instead.
                if (dstDoc.getLists().getListByListId(listId) != null)
                {
                    List currentList;
                    // A newly copied list already exists for this ID, retrieve the stored list and use it on
                    // the current paragraph.
                    if (newLists.containsKey(listId))
                    {
                        currentList = (List)newLists.get(listId);
                    }
                    else
                    {
                        // Add a copy of this list to the document and store it for later reference.
                        currentList = srcDoc.getLists().addCopy(para.getListFormat().getList());
                        newLists.put(listId, currentList);
                    }

                    // Set the list of this paragraph  to the copied list.
                    para.getListFormat().setList(currentList);
                }
            }
        }

        // Append the source document to end of the destination document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES);

        // Save the combined document to disk.
        dstDoc.save(gDataDir + "TestFile.ListUseDestinationStyles Out.docx");

        System.out.println("Documents appended successfully.");
    }
}