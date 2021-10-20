package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;
import org.testng.Assert;

public class ProtectDocument {

	public static final String dataDir = Utils.getSharedDataDir(ProtectDocument.class) + "Document/";

	public static void main(String[] args) throws Exception {
		// Protecting a Document
		PasswordProtection();
		AllowOnlyFormFieldsProtect();

		// Unprotecting a Document
		RemoveDocumentProtection();

		// Getting the Protection Type
		getTheProtectionType();

		UnrestrictedEditableRegions();
		UnrestrictedSection();
		ReadOnlyProtection();
		RemoveReadOnlyRestriction();
	}

	public static void PasswordProtection() throws Exception {
		//ExStart:PasswordProtection
		// Create a new document and protect it with a password.
		Document doc = new Document();

		// Apply Document Protection.
		doc.protect(ProtectionType.NO_PROTECTION, "password");

		// Save the document.
		doc.save(dataDir + "ProtectDocument.PasswordProtection.docx");
		//ExEnd:PasswordProtection
	}

	public static void AllowOnlyFormFieldsProtect() throws Exception {
		//ExStart:AllowOnlyFormFieldsProtect
		// Insert two sections with some text.
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		builder.writeln("Text added to a document.");

		// A document protection only works when document protection is turned and only editing in form fields is allowed.
		doc.protect(ProtectionType.ALLOW_ONLY_FORM_FIELDS, "password");

		// Save the protected document.
		doc.save(dataDir + "ProtectDocument.AllowOnlyFormFieldsProtect.docx");
		//ExEnd:AllowOnlyFormFieldsProtect
	}

	public static void RemoveDocumentProtection() throws Exception {
		//ExStart:RemoveDocumentProtection
		// Create a new document and protect it with a password.
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		builder.writeln("Text added to a document.");

		// Documents can have protection removed either with no password, or with the correct password.
		doc.unprotect();
		doc.protect(ProtectionType.READ_ONLY, "newPassword");
		doc.unprotect("newPassword");

		doc.save(dataDir + "ProtectDocument.RemoveDocumentProtection.docx");
		//ExEnd:RemoveDocumentProtection
	}

	public static void UnrestrictedEditableRegions() throws Exception {
		//ExStart:UnrestrictedEditableRegions
		// Upload a document and make it as read-only.
		Document doc = new Document(dataDir + "Document.docx");
		DocumentBuilder builder = new DocumentBuilder(doc);
		doc.protect(ProtectionType.READ_ONLY, "MyPassword");
		builder.writeln("Hello world! Since we have set the document's protection level to read-only, " + "we cannot edit this paragraph without the password.");

		// Start an editable range.
		EditableRangeStart edRangeStart = builder.startEditableRange();

		// An EditableRange object is created for the EditableRangeStart that we just made.
		EditableRange editableRange = edRangeStart.getEditableRange();

		// Put something inside the editable range.
		builder.writeln("Paragraph inside first editable range");

		// An editable range is well-formed if it has a start and an end.
		EditableRangeEnd edRangeEnd = builder.endEditableRange();

		builder.writeln("This paragraph is outside any editable ranges, and cannot be edited.");
		doc.save(dataDir + "ProtectDocument.UnrestrictedEditableRegions.docx");
		//ExEnd:UnrestrictedEditableRegions
	}

	public static void UnrestrictedSection() throws Exception {
		//ExStart:UnrestrictedSection
		// Insert two sections with some text.
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		builder.writeln("Section 1. Unprotected.");
		builder.insertBreak(BreakType.SECTION_BREAK_CONTINUOUS);
		builder.writeln("Section 2. Protected.");

		// Section protection only works when document protection is turned and only editing in form fields is allowed.
		doc.protect(ProtectionType.ALLOW_ONLY_FORM_FIELDS, "password");

		// By default, all sections are protected, but we can selectively turn protection off.
		doc.getSections().get(0).setProtectedForForms(false);
		doc.save(dataDir + "Section.Protect.docx");

		doc = new Document(dataDir + "Section.Protect.docx");
		Assert.assertFalse(doc.getSections().get(0).getProtectedForForms());
		Assert.assertTrue(doc.getSections().get(1).getProtectedForForms());
		//ExEnd:UnrestrictedSection
	}

	public static void getTheProtectionType() throws Exception {
		// ExStart:getTheProtectionType
		Document doc = new Document(dataDir + "Document.doc");
		int protectionType = doc.getProtectionType();
		// ExEnd:getTheProtectionType
	}

	public static void ReadOnlyProtection() throws Exception {
		//ExStart:ReadOnlyProtection
		// Create a document.
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		// Add text.
		builder.write("Open document as read-only");

		// Enter a password that's up to 15 characters long.
		doc.getWriteProtection().setPassword("MyPassword");

		// Make the document as read-only.
		doc.getWriteProtection().setReadOnlyRecommended(true);

		// Apply write protection as read-only.
		doc.protect(ProtectionType.READ_ONLY);
		doc.save(dataDir + "ProtectDocument.ReadOnlyProtection.docx");
		//ExEnd:ReadOnlyProtection
	}

	public static void RemoveReadOnlyRestriction() throws Exception {
		//ExStart:RemoveReadOnlyRestriction
		Document doc = new Document();

		// Enter a password that's up to 15 characters long.
		doc.getWriteProtection().setPassword("MyPassword");

		// Remove the read-only option.
		doc.getWriteProtection().setReadOnlyRecommended(false);

		// Apply write protection without any protection.
		doc.protect(ProtectionType.NO_PROTECTION);
		doc.save(dataDir + "ProtectDocument.RemoveReadOnlyRestriction.docx");
		//ExEnd:RemoveReadOnlyRestriction
	}

}
