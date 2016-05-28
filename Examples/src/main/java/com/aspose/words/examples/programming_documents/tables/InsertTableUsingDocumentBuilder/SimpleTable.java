/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.InsertTableUsingDocumentBuilder;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class SimpleTable
{
    public static void main(String[] args) throws Exception
    {
		//ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(SimpleTable.class);

		// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		// We call this method to start building the table.
		builder.startTable();
		builder.insertCell();
		builder.write("Row 1, Cell 1 Content.");
		// Build the second cell
		builder.insertCell();
		builder.write("Row 1, Cell 2 Content.");
		// Call the following method to end the row and start a new row.
		builder.endRow();

		// Build the first cell of the second row.
		builder.insertCell();
		builder.write("Row 2, Cell 1 Content");

		// Build the second cell.
		builder.insertCell();
		builder.write("Row 2, Cell 2 Content.");
		builder.endRow();
		// Signal that we have finished building the table.
		builder.endTable();
		// Save the document to disk.
		doc.save(dataDir + "output.doc");
		//ExEnd:1
    }
}