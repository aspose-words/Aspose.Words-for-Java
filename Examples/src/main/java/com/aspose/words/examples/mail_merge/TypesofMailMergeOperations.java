package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;
import com.aspose.words.net.System.Data.DataSet;
import com.aspose.words.net.System.Data.DataTable;

public class TypesofMailMergeOperations {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(TypesofMailMergeOperations.class) + "MailMerge/";
		ExecuteSimpleMailMerge(dataDir);
		MailMergeWithRegions(dataDir);
		NestedMailMerge(dataDir);
	}

	private static void ExecuteSimpleMailMerge(String dataDir) throws Exception {
		// ExStart:ExecuteSimpleMailMerge
		// Include the code for our template.
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		// Create Merge Fields.
		builder.insertField(" MERGEFIELD CustomerName ");
		builder.insertParagraph();
		builder.insertField(" MERGEFIELD Item ");
		builder.insertParagraph();
		builder.insertField(" MERGEFIELD Quantity ");

		builder.getDocument().save(dataDir + "MailMerge.TestTemplate.docx");

		// Fill the fields in the document with user data.
		doc.getMailMerge().execute(new String[] { "CustomerName", "Item", "Quantity" },
				new Object[] { "John Doe", "Hawaiian", "2" });

		builder.getDocument().save(dataDir + "MailMerge.Simple.docx");
		// ExEnd:ExecuteSimpleMailMerge
	}

	public static void MailMergeWithRegions(String dataDir) throws Exception {
		// ExStart: MailMergeWithRegions
		// For complete examples and data files, please go to
		// https://github.com/aspose-words/Aspose.Words-for-java
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		// The start point of mail merge with regions the dataset.
		builder.insertField(" MERGEFIELD TableStart:Customers");
		// Data from rows of the "CustomerName" column of the "Customers" table will go
		// in this MERGEFIELD.
		builder.write("Orders for ");
		builder.insertField(" MERGEFIELD CustomerName");
		builder.write(":");

		// Create column headers
		builder.startTable();
		builder.insertCell();
		builder.write("Item");
		builder.insertCell();
		builder.write("Quantity");
		builder.endRow();

		// We have a second data table called "Orders", which has a many-to-one
		// relationship with "Customers"
		// picking up rows with the same CustomerID value.
		builder.insertCell();
		builder.insertField(" MERGEFIELD TableStart:Orders");
		builder.insertField(" MERGEFIELD ItemName");
		builder.insertCell();
		builder.insertField(" MERGEFIELD Quantity");
		builder.insertField(" MERGEFIELD TableEnd:Orders");
		builder.endTable();

		// The end point of mail merge with regions.
		builder.insertField(" MERGEFIELD TableEnd:Customers");

		// Pass our dataset to perform mail merge with regions.
		DataSet customersAndOrders = CreateDataSet();
		doc.getMailMerge().executeWithRegions(customersAndOrders);

		// Save the result
		doc.save(dataDir + "MailMerge.ExecuteWithRegions.docx");
		// ExEnd: MailMergeWithRegions
	}

	//ExStart:CreateDataSet
	private static DataSet CreateDataSet() {
		// Create the customers table.
		DataTable tableCustomers = new DataTable("Customers");
		tableCustomers.getColumns().add("CustomerID");
		tableCustomers.getColumns().add("CustomerName");
		tableCustomers.getRows().add(new Object[] { 1, "John Doe" });
		tableCustomers.getRows().add(new Object[] { 2, "Jane Doe" });

		// Create the orders table.
		DataTable tableOrders = new DataTable("Orders");
		tableOrders.getColumns().add("CustomerID");
		tableOrders.getColumns().add("ItemName");
		tableOrders.getColumns().add("Quantity");
		tableOrders.getRows().add(new Object[] { 1, "Hawaiian", 2 });
		tableOrders.getRows().add(new Object[] { 2, "Pepperoni", 1 });
		tableOrders.getRows().add(new Object[] { 2, "Chicago", 1 });

		// Add both tables to a data set.
		DataSet dataSet = new DataSet();
		dataSet.getTables().add(tableCustomers);
		dataSet.getTables().add(tableOrders);

		// The "CustomerID" column, also the primary key of the customers table is the
		// foreign key for the Orders table.
		dataSet.getRelations().add(tableCustomers.getColumns().get("CustomerID"),
				tableOrders.getColumns().get("CustomerID"));

		return dataSet;
	}
	//ExEnd:CreateDataSet
	
	private static void NestedMailMerge(String dataDir) throws Exception {
		// ExStart: NestedMailMerge
		// Create the Dataset and read the XML.
		DataSet pizzaDs = new DataSet();
			 
		// The Datatable.TableNames and the DataSet.Relations are defined implicitly by .NET through ReadXml.
		pizzaDs.readXml(dataDir + "CustomerData.xml");
		String fileName = "Invoice Template.doc";

		// Open the template document.
		Document doc = new Document(dataDir + fileName);
			 
		// Trim trailing and leading whitespaces mail merge values.
		doc.getMailMerge().setTrimWhitespaces(false);
			 
		// Execute the nested mail merge with regions.
		doc.getMailMerge().executeWithRegions(pizzaDs);

		// Save the output to file.
		doc.save(dataDir + fileName);
		// ExEnd: NestedMailMerge
		System.out.println("\nMail merge performed with nested data successfully.\nFile saved at " + dataDir);
	}
}
