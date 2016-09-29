package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class ExecuteSimpleMailMerge {

	private static final String dataDir = Utils.getSharedDataDir(ExecuteSimpleMailMerge.class) + "MailMerge/";

	public static void main(String[] args) throws Exception {
		// Open an existing document.
		Document doc = new Document(dataDir + "MailMerge.ExecuteArray.doc");

		// Trim trailing and leading whitespaces mail merge values
		doc.getMailMerge().setTrimWhitespaces(false);

		// Fill the fields in the document with user data.
		doc.getMailMerge().execute(new String[] { "FullName", "Company", "Address", "Address2", "City" }, 
				new Object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

		doc.save(dataDir + "MailMerge.ExecuteArray_Out.doc");
	}
}