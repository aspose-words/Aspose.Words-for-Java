/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables;

import com.aspose.words.AutoFitBehavior;
import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;


public class AutoFitTableToFixedWidthColumns
{
    public static void main(String[] args) throws Exception
    {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(AutoFitTableToFixedWidthColumns.class);

		Document doc = new Document(dataDir + "TestFile.doc");
		Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);

		// Disable autofitting on this table.
		table.autoFit(AutoFitBehavior.FIXED_COLUMN_WIDTHS);

		// Save the document to disk.
		doc.save(dataDir + "TestFile.FixedWidth Out.doc");

		System.out.println("Table auto fit to fixed width columns successfully.");
    }
}