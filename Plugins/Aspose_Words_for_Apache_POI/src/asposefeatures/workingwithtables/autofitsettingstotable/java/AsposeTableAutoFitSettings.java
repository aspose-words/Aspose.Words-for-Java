package asposefeatures.workingwithtables.autofitsettingstotable.java;

import com.aspose.words.AutoFitBehavior;
import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.Table;

public class AsposeTableAutoFitSettings
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/asposefeatures/workingwithtables/autofitsettingstotable/data/";

		// Open the document
		Document doc = new Document(dataPath + "tableDoc.doc");
		
		Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);
		// Autofit the first table to the page width.
		table.autoFit(AutoFitBehavior.AUTO_FIT_TO_WINDOW);

		Table table2 = (Table)doc.getChild(NodeType.TABLE, 1, true);
		// Auto fit the table to the cell contents
		table2.autoFit(AutoFitBehavior.AUTO_FIT_TO_CONTENTS);

		// Save the document to disk.
		doc.save(dataPath + "AsposeAutoFitTable_Out.doc");
		
		System.out.println("Process Completed Successfully");
	}
}