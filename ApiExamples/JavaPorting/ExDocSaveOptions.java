// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.DocSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.ms.System.IO.Directory;


@Test
public class ExDocSaveOptions extends ApiExampleBase
{
    @Test
    public void saveAsDoc() throws Exception
    {
        //ExStart
        //ExFor:DocSaveOptions
        //ExFor:DocSaveOptions.#ctor
        //ExFor:DocSaveOptions.#ctor(SaveFormat)
        //ExFor:DocSaveOptions.Password
        //ExFor:DocSaveOptions.SaveFormat
        //ExFor:DocSaveOptions.SaveRoutingSlip
        //ExSummary:Shows how to set save options for classic Microsoft Word document versions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Hello world!");

        // DocSaveOptions only applies to Doc and Dot save formats
        DocSaveOptions options = new DocSaveOptions(SaveFormat.DOC);

        // Set a password with which the document will be encrypted, and which will be required to open it
        options.setPassword("MyPassword");

        // If the document contains a routing slip, we can preserve it while saving by setting this flag to true
        options.setSaveRoutingSlip(true);

        doc.save(getArtifactsDir() + "DocSaveOptions.SaveAsDoc.doc", options);          
        //ExEnd
    }

    @Test
    public void tempFolder() throws Exception
    {
        //ExStart
        //ExFor:SaveOptions.TempFolder
        //ExSummary:Shows how to save a document using temporary files.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // We can use a SaveOptions object to set the saving method of a document from a MemoryStream to temporary files
        // While saving, the files will briefly pop up in the folder we set as the TempFolder attribute below
        // Doing this will free up space in the memory that the stream would usually occupy
        DocSaveOptions options = new DocSaveOptions();
        options.setTempFolder(getArtifactsDir() + "TempFiles");

        // Ensure that the directory exists and save
        Directory.createDirectory(options.getTempFolder());

        doc.save(getArtifactsDir() + "DocSaveOptions.TempFolder.doc", options);
        //ExEnd
    }

    @Test
    public void pictureBullets() throws Exception
    {
        //ExStart
        //ExFor:DocSaveOptions.SavePictureBullet
        //ExSummary:Shows how to remove PictureBullet data from the document.
        Document doc = new Document(getMyDir() + "Document.PictureBullets.docx");

        // Word 97 cannot work correctly with PictureBullet data
        // To remove PictureBullet data, set the option to "false"
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.DOC);
        saveOptions.setSavePictureBullet(false);

        doc.save(getArtifactsDir() + "Document.PictureBullets.doc", saveOptions);
        //ExEnd
    }
}
