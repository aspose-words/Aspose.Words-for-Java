package DocsExamples.Programming_with_Documents.Protect_or_Encrypt_Document;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.ProtectionType;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.EditableRangeStart;
import com.aspose.words.EditableRange;
import com.aspose.words.EditableRangeEnd;
import com.aspose.words.BreakType;
import org.testng.Assert;


class DocumentProtection extends DocsExamplesBase
{
    @Test
    public void passwordProtection() throws Exception
    {
        //ExStart:PasswordProtection
        Document doc = new Document();

        // Apply document protection.
        doc.protect(ProtectionType.NO_PROTECTION, "password");

        doc.save(getArtifactsDir() + "DocumentProtection.PasswordProtection.docx");
        //ExEnd:PasswordProtection
    }

    @Test
    public void allowOnlyFormFieldsProtect() throws Exception
    {
        //ExStart:AllowOnlyFormFieldsProtect
        // Insert two sections with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Text added to a document.");

        // A document protection only works when document protection is turned and only editing in form fields is allowed.
        doc.protect(ProtectionType.ALLOW_ONLY_FORM_FIELDS, "password");

        // Save the protected document.
        doc.save(getArtifactsDir() + "DocumentProtection.AllowOnlyFormFieldsProtect.docx");
        //ExEnd:AllowOnlyFormFieldsProtect
    }

    @Test
    public void removeDocumentProtection() throws Exception
    {
        //ExStart:RemoveDocumentProtection
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Text added to a document.");

        // Documents can have protection removed either with no password, or with the correct password.
        doc.unprotect();
        doc.protect(ProtectionType.READ_ONLY, "newPassword");
        doc.unprotect("newPassword");

        doc.save(getArtifactsDir() + "DocumentProtection.RemoveDocumentProtection.docx");
        //ExEnd:RemoveDocumentProtection
    }

    @Test
    public void unrestrictedEditableRegions() throws Exception
    {
        //ExStart:UnrestrictedEditableRegions
        // Upload a document and make it as read-only.
        Document doc = new Document(getMyDir() + "Document.docx");
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

        doc.save(getArtifactsDir() + "DocumentProtection.UnrestrictedEditableRegions.docx");
        //ExEnd:UnrestrictedEditableRegions
    }

    @Test
    public void unrestrictedSection() throws Exception
    {
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
        doc.save(getArtifactsDir() + "DocumentProtection.UnrestrictedSection.docx");

        doc = new Document(getArtifactsDir() + "DocumentProtection.UnrestrictedSection.docx");
        Assert.assertFalse(doc.getSections().get(0).getProtectedForForms());
        Assert.assertTrue(doc.getSections().get(1).getProtectedForForms());
        //ExEnd:UnrestrictedSection
    }

    @Test
    public void getProtectionType() throws Exception
    {
        //ExStart:GetProtectionType
        Document doc = new Document(getMyDir() + "Document.docx");
        /*ProtectionType*/int protectionType = doc.getProtectionType();
        //ExEnd:GetProtectionType
    }

    @Test
    public void readOnlyProtection() throws Exception
    {
        //ExStart:ReadOnlyProtection
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Open document as read-only");

        // Enter a password that's up to 15 characters long.
        doc.getWriteProtection().setPassword("MyPassword");

        // Make the document as read-only.
        doc.getWriteProtection().setReadOnlyRecommended(true);

        // Apply write protection as read-only.
        doc.protect(ProtectionType.READ_ONLY);
        doc.save(getArtifactsDir() + "DocumentProtection.ReadOnlyProtection.docx");
        //ExEnd:ReadOnlyProtection
    }

    @Test
    public void removeReadOnlyRestriction() throws Exception
    {
        //ExStart:RemoveReadOnlyRestriction
        Document doc = new Document();
        
        // Enter a password that's up to 15 characters long.
        doc.getWriteProtection().setPassword("MyPassword");

        // Remove the read-only option.
        doc.getWriteProtection().setReadOnlyRecommended(false);

        // Apply write protection without any protection.
        doc.protect(ProtectionType.NO_PROTECTION);
        doc.save(getArtifactsDir() + "DocumentProtection.RemoveReadOnlyRestriction.docx");
        //ExEnd:RemoveReadOnlyRestriction
    }
}
