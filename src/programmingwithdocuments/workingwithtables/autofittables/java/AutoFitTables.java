/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package programmingwithdocuments.workingwithtables.autofittables.java;

import java.io.File;
import java.net.URI;

import com.aspose.words.*;


public class AutoFitTables
{
    public static void main(String[] args) throws Exception
    {
            // The path to the documents directory.
        String dataDir = "src/programmingwithdocuments/workingwithtables/autofittables/data/";

        // Demonstrate autofitting a table to the window.
        autoFitTableToWindow(dataDir);

        // Demonstrate autofitting a table to its contents.
        autoFitTableToContents(dataDir);

        // Demonstrate autofitting a table to fixed column widths.
        autoFitTableToFixedColumnWidths(dataDir);
    }

    public static void autoFitTableToWindow(String dataDir) throws Exception
    {
		 //ExStart
		 //ExFor:Table.AutoFit
		 //ExFor:AutoFitBehavior
		 //ExId:FitTableToPageWidth
		 //ExSummary:Autofits a table to fit the page width.
		 // Open the document
		 Document doc = new Document(dataDir + "TestFile.doc");
		 Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);

		 // Autofit the first table to the page width.
		 table.autoFit(AutoFitBehavior.AUTO_FIT_TO_WINDOW);

		 // Save the document to disk.
		 doc.save(dataDir + "TestFile.AutoFitToWindow Out.doc");
		 //ExEnd

		 assert(doc.getFirstSection().getBody().getTables().get(0).getPreferredWidth().getType() == PreferredWidthType.PERCENT) : "PreferredWidth type is not percent";
		 assert(doc.getFirstSection().getBody().getTables().get(0).getPreferredWidth().getValue() == 100) : "PreferredWidth value is different than 100";

	}

	public static void autoFitTableToContents(String dataDir) throws Exception
	{
		  //ExStart
		  //ExFor:Table.AutoFit
		  //ExFor:AutoFitBehavior
		  //ExId:FitTableToContents
		  //ExSummary:Autofits a table in the document to its contents.
		  // Open the document
		  Document doc = new Document(dataDir + "TestFile.doc");
		  Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);

		  // Auto fit the table to the cell contents
		  table.autoFit(AutoFitBehavior.AUTO_FIT_TO_CONTENTS);

		  // Save the document to disk.
		  doc.save(dataDir + "TestFile.AutoFitToContents Out.doc");
		  //ExEnd

		  assert(doc.getFirstSection().getBody().getTables().get(0).getPreferredWidth().getType() == PreferredWidthType.AUTO) : "PreferredWidth type is not auto";
		  assert(doc.getFirstSection().getBody().getTables().get(0).getFirstRow().getFirstCell().getCellFormat().getPreferredWidth().getType() == PreferredWidthType.AUTO) : "PrefferedWidth on cell is not auto";
          assert(doc.getFirstSection().getBody().getTables().get(0).getFirstRow().getFirstCell().getCellFormat().getPreferredWidth().getValue() == 0) : "PreferredWidth value is not 0";
	}

	public static void autoFitTableToFixedColumnWidths(String dataDir) throws Exception
	{
		 //ExStart
		 //ExFor:Table.AutoFit
		 //ExFor:AutoFitBehavior
		 //ExId:DisableAutoFitAndUseFixedWidths
		 //ExSummary:Disables autofitting and enables fixed widths for the specified table.
		 // Open the document
		 Document doc = new Document(dataDir + "TestFile.doc");
		 Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);

		 // Disable autofitting on this table.
		 table.autoFit(AutoFitBehavior.FIXED_COLUMN_WIDTHS);

		 // Save the document to disk.
		 doc.save(dataDir + "TestFile.FixedWidth Out.doc");
		 //ExEnd

		 assert(doc.getFirstSection().getBody().getTables().get(0).getPreferredWidth().getType() == PreferredWidthType.AUTO) : "PreferredWidth type is not auto";
		 assert(doc.getFirstSection().getBody().getTables().get(0).getPreferredWidth().getValue() == 0) : "PreferredWidth value is not 0";
         assert(doc.getFirstSection().getBody().getTables().get(0).getFirstRow().getFirstCell().getCellFormat().getWidth() == 69.2) : "Cell width is not correct.";
	}
}