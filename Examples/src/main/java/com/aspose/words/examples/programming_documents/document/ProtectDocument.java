package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.ProtectionType;
import com.aspose.words.examples.Utils;

public class ProtectDocument {

	public static final String dataDir = Utils.getSharedDataDir(ProtectDocument.class) + "Document/";

	public static void main(String[] args) throws Exception {
		// Protecting a Document
		protectADocument();

		// Unprotecting a Document
		unprotectADocument();

		// Getting the Protection Type
		getTheProtectionType();
	}

	public static void protectADocument() throws Exception {
		// ExStart:protectADocument
		Document doc = new Document();
		doc.protect(ProtectionType.ALLOW_ONLY_FORM_FIELDS, "password");
		// ExEnd:protectADocument
	}

	public static void unprotectADocument() throws Exception {
		// ExStart:unprotectADocument
		Document doc = new Document();
		doc.unprotect();
		// ExEnd:unprotectADocument
	}

	public static void getTheProtectionType() throws Exception {
		// ExStart:getTheProtectionType
		Document doc = new Document(dataDir + "Document.doc");
		int protectionType = doc.getProtectionType();
		// ExEnd:getTheProtectionType
	}

}
