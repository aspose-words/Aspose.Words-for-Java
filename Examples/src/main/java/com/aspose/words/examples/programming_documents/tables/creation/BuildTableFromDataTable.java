package com.aspose.words.examples.programming_documents.tables.creation;

import java.sql.ResultSetMetaData;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Orientation;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.Table;
import com.aspose.words.TableStyleOptions;
import com.aspose.words.examples.Utils;
import com.aspose.words.net.System.Data.DataTable;

public class BuildTableFromDataTable {
	
	private static final String dataDir = Utils.getSharedDataDir(BuildTableFromDataTable.class) + "Tables/";
	
	public static void main(String[] args) throws Exception {
		// Create a new document.
		Document doc = new Document();

		// We can position where we want the table to be inserted and also specify any extra formatting to be
		// applied onto the table as well.
		DocumentBuilder builder = new DocumentBuilder(doc);

		// We want to rotate the page landscape as we expect a wide table.
		doc.getFirstSection().getPageSetup().setOrientation(Orientation.LANDSCAPE);

		// Retrieve the data from our data source which is stored as a DataTable.
		DataTable dataTable = null; //getEmployees(databaseDir);

		// Build a table in the document from the data contained in the DataTable.
		Table table = importTableFromDataTable(builder, dataTable, true);

		// We can apply a table style as a very quick way to apply formatting to the entire table.
		table.setStyleIdentifier(StyleIdentifier.MEDIUM_LIST_2_ACCENT_1);
		table.setStyleOptions(TableStyleOptions.FIRST_ROW | TableStyleOptions.ROW_BANDS | TableStyleOptions.LAST_COLUMN);

		// For our table we want to remove the heading for the image column.
		table.getFirstRow().getLastCell().removeAllChildren();

		doc.save(dataDir + "Table.FromDataTable_Out.docx");
	}

	/*
	 * Imports the content from the specified DataTable into a new Aspose.Words
	 * Table object. The table is inserted at the current position of the
	 * document builder and using the current builder's formatting if any is
	 * defined.
	 */
	public static Table importTableFromDataTable(DocumentBuilder builder, DataTable dataTable, boolean importColumnHeadings) throws Exception {
		Table table = builder.startTable();

		ResultSetMetaData metaData = dataTable.getResultSet().getMetaData();
		int numColumns = metaData.getColumnCount();

		// Check if the names of the columns from the data source are to be included in a header row.
		if (importColumnHeadings) {
			// Store the original values of these properties before changing them.
			boolean boldValue = builder.getFont().getBold();
			int paragraphAlignmentValue = builder.getParagraphFormat().getAlignment();

			// Format the heading row with the appropriate properties.
			builder.getFont().setBold(true);
			builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

			// Create a new row and insert the name of each column into the first row of the table.
			for (int i = 1; i < numColumns + 1; i++) {
				builder.insertCell();
				builder.writeln(metaData.getColumnName(i));
			}

			builder.endRow();

			// Restore the original formatting.
			builder.getFont().setBold(boldValue);
			builder.getParagraphFormat().setAlignment(paragraphAlignmentValue);
		}

		// Iterate through all rows and then columns of the data.
		while (dataTable.getResultSet().next()) {
			for (int i = 1; i < numColumns + 1; i++) {
				// Insert a new cell for each object.
				builder.insertCell();

				// Retrieve the current record.
				Object item = dataTable.getResultSet().getObject(metaData.getColumnName(i));
				// This is name of the data type.
				String typeName = item.getClass().getSimpleName();

				if (typeName.equals("byte[]")) {
					// Assume a byte array is an image. Other data types can be added here.
					builder.insertImage((byte[]) item, 50, 50);
				} else if (typeName.equals("Timestamp")) {
					// Define a custom format for dates and times.
					builder.write(new SimpleDateFormat("MMMM d, yyyy").format((Timestamp) item));
				} else {
					// By default any other item will be inserted as text.
					builder.write(item.toString());
				}
			}

			// After we insert all the data from the current record we can end the table row.
			builder.endRow();
		}

		// We have finished inserting all the data from the DataTable, we can end the table.
		builder.endTable();
		return table;
	}
}
