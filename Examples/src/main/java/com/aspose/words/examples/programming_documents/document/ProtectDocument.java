package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.ProtectionType;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.programming_documents.document.properties.AccessingDocumentProperties;

public class ProtectDocument {
	
	public static final String dataDir = Utils.getSharedDataDir(AccessingDocumentProperties.class) + "Document/";
	
	public static void main(String[] args) throws Exception {
		// Protecting a Document
		protectADocument();
		
		// Unprotecting a Document
		unprotectADocument();
		
		// Getting the Protection Type
		getTheProtectionType();
	}
	
	public static void protectADocument() throws Exception {
		Document doc = new Document();
		doc.protect(ProtectionType.ALLOW_ONLY_FORM_FIELDS, "password");
	}
	
	public static void unprotectADocument() throws Exception {
		Document doc = new Document();
		doc.unprotect();
	}
	
	public static void getTheProtectionType() throws Exception {
		Document doc = new Document(dataDir + "Document.doc");
		int protectionType = doc.getProtectionType();
	}

}
