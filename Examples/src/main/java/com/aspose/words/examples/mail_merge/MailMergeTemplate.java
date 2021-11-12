package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.TextFormFieldType;
import com.aspose.words.BreakType;

public class MailMergeTemplate {
	// ExStart: CreateMailMergeTemplate
	public static Document CreateMailMergeTemplate() throws Exception	{
	    DocumentBuilder builder = new DocumentBuilder();
	    
	 // Insert a text input field the unique name of this field is "Hello", the other parameters define
	    // what type of FormField it is, the format of the text, the field result and the maximum text length (0 = no limit)
	    builder.insertTextInput("TextInput", TextFormFieldType.REGULAR, "", "Hello", 0);
	    builder.insertField("MERGEFIELD CustomerFirstName \\* MERGEFORMAT");

	    builder.insertTextInput("TextInput1", TextFormFieldType.REGULAR, "", " ", 0);
	    builder.insertField("MERGEFIELD CustomerLastName \\* MERGEFORMAT");

	    builder.insertTextInput("TextInput1", TextFormFieldType.REGULAR, "", " , ", 0);

	    // Inserts a paragraph break into the document
	    builder.insertParagraph();

	    // Insert mail body
	    builder.insertTextInput("TextInput", TextFormFieldType.REGULAR, "", "Thanks for purchasing our ", 0);
	    builder.insertField("MERGEFIELD ProductName \\* MERGEFORMAT");

	    builder.insertTextInput("TextInput", TextFormFieldType.REGULAR, "", ", please download your Invoice at ",
	        0);
	    builder.insertField("MERGEFIELD InvoiceURL \\* MERGEFORMAT");

	    builder.insertTextInput("TextInput", TextFormFieldType.REGULAR, "",
	        ". If you have any questions please call ", 0);
	    builder.insertField("MERGEFIELD Supportphone \\* MERGEFORMAT");

	    builder.insertTextInput("TextInput", TextFormFieldType.REGULAR, "", ", or email us at ", 0);
	    builder.insertField("MERGEFIELD SupportEmail \\* MERGEFORMAT");

	    builder.insertTextInput("TextInput", TextFormFieldType.REGULAR, "", ".", 0);

	    builder.insertParagraph();

	    // Insert mail ending
	    builder.insertTextInput("TextInput", TextFormFieldType.REGULAR, "", "Best regards,", 0);
	    builder.insertBreak(BreakType.LINE_BREAK);
	    builder.insertField("MERGEFIELD EmployeeFullname \\* MERGEFORMAT");

	    builder.insertTextInput("TextInput1", TextFormFieldType.REGULAR, "", " ", 0);
	    builder.insertField("MERGEFIELD EmployeeDepartment \\* MERGEFORMAT");
	    
	    return builder.getDocument();
	}
	// ExEnd: CreateMailMergeTemplate
}
