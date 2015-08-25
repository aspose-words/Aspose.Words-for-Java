package asposefeatures.workingwithtables.splittables.java;

import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.Row;
import com.aspose.words.Table;

public class AsposeSplittingTables
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/asposefeatures/workingwithtables/splittables/data/";
		
		// Load the document.
		Document doc = new Document(dataPath + "tableDoc.doc");

		// Get the first table in the document.
		Table firstTable = (Table)doc.getChild(NodeType.TABLE, 0, true);

		// We will split the table at the third row (inclusive).
		Row row = firstTable.getRows().get(2);

		// Create a new container for the split table.
		Table table = (Table)firstTable.deepClone(false);

		// Insert the container after the original.
		firstTable.getParentNode().insertAfter(table, firstTable);

		// Add a buffer paragraph to ensure the tables stay apart.
		firstTable.getParentNode().insertAfter(new Paragraph(doc), firstTable);

		Row currentRow;

		do
		{
		    currentRow = firstTable.getLastRow();
		    table.prependChild(currentRow);
		}
		while (currentRow != row);

		doc.save(dataPath + "AsposeSplitTable.doc");
		
		System.out.println("Done.");
	}
}