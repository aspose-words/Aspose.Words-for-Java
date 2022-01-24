package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.FileInputStream;

public class ExPlainTextDocument extends ApiExampleBase {
    @Test
    public void load() throws Exception {
        //ExStart
        //ExFor:PlainTextDocument
        //ExFor:PlainTextDocument.#ctor(String)
        //ExFor:PlainTextDocument.Text
        //ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        doc.save(getArtifactsDir() + "PlainTextDocument.Load.docx");

        PlainTextDocument plaintext = new PlainTextDocument(getArtifactsDir() + "PlainTextDocument.Load.docx");

        Assert.assertEquals("Hello world!", plaintext.getText().trim());
        //ExEnd
    }

    @Test
    public void loadFromStream() throws Exception {
        //ExStart
        //ExFor:PlainTextDocument.#ctor(Stream)
        //ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext using stream.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");
        doc.save(getArtifactsDir() + "PlainTextDocument.LoadFromStream.docx");

        try (FileInputStream stream = new FileInputStream(getArtifactsDir() + "PlainTextDocument.LoadFromStream.docx")) {
            PlainTextDocument plaintext = new PlainTextDocument(stream);

            Assert.assertEquals("Hello world!", plaintext.getText().trim());
        }
        //ExEnd
    }

    @Test
    public void loadEncrypted() throws Exception {
        //ExStart
        //ExFor:PlainTextDocument.#ctor(String, LoadOptions)
        //ExSummary:Shows how to load the contents of an encrypted Microsoft Word document in plaintext.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.setPassword("MyPassword");

        doc.save(getArtifactsDir() + "PlainTextDocument.LoadEncrypted.docx", saveOptions);

        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setPassword("MyPassword");

        PlainTextDocument plaintext = new PlainTextDocument(getArtifactsDir() + "PlainTextDocument.LoadEncrypted.docx", loadOptions);

        Assert.assertEquals("Hello world!", plaintext.getText().trim());
        //ExEnd
    }

    @Test
    public void loadEncryptedUsingStream() throws Exception {
        //ExStart
        //ExFor:PlainTextDocument.#ctor(Stream, LoadOptions)
        //ExSummary:Shows how to load the contents of an encrypted Microsoft Word document in plaintext using stream.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.setPassword("MyPassword");

        doc.save(getArtifactsDir() + "PlainTextDocument.LoadFromStreamWithOptions.docx", saveOptions);

        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setPassword("MyPassword");

        try (FileInputStream stream = new FileInputStream(getArtifactsDir() + "PlainTextDocument.LoadFromStreamWithOptions.docx")) {
            PlainTextDocument plaintext = new PlainTextDocument(stream, loadOptions);

            Assert.assertEquals("Hello world!", plaintext.getText().trim());
        }
        //ExEnd
    }

    @Test
    public void builtInProperties() throws Exception {
        //ExStart
        //ExFor:PlainTextDocument.BuiltInDocumentProperties
        //ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext and then access the original document's built-in properties.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");
        doc.getBuiltInDocumentProperties().setAuthor("John Doe");

        doc.save(getArtifactsDir() + "PlainTextDocument.BuiltInProperties.docx");

        PlainTextDocument plaintext = new PlainTextDocument(getArtifactsDir() + "PlainTextDocument.BuiltInProperties.docx");

        Assert.assertEquals("Hello world!", plaintext.getText().trim());
        Assert.assertEquals("John Doe", plaintext.getBuiltInDocumentProperties().getAuthor());
        //ExEnd
    }

    @Test
    public void customDocumentProperties() throws Exception {
        //ExStart
        //ExFor:PlainTextDocument.CustomDocumentProperties
        //ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext and then access the original document's custom properties.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");
        doc.getCustomDocumentProperties().add("Location of writing", "123 Main St, London, UK");

        doc.save(getArtifactsDir() + "PlainTextDocument.CustomDocumentProperties.docx");

        PlainTextDocument plaintext = new PlainTextDocument(getArtifactsDir() + "PlainTextDocument.CustomDocumentProperties.docx");

        Assert.assertEquals("Hello world!", plaintext.getText().trim());
        Assert.assertEquals("123 Main St, London, UK", plaintext.getCustomDocumentProperties().get("Location of writing").getValue());
        //ExEnd
    }
}

