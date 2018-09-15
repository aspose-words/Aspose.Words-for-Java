package com.aspose.words.examples.loading_saving;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class Load_Options {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(Load_Options.class);

        loadOptionsUpdateDirtyFields(dataDir);
        loadAndSaveEncryptedODT(dataDir);
        verifyODTdocument(dataDir);
        convertShapeToOfficeMath(dataDir);
        annotationsAtBlockLevel(dataDir);
    }

    public static void loadOptionsUpdateDirtyFields(String dataDir) throws Exception
    {
        // ExStart:LoadOptionsUpdateDirtyFields
        LoadOptions lo = new LoadOptions();
        //Update the fields with the dirty attribute
        lo.setUpdateDirtyFields(true);
        //Load the Word document
        Document doc = new Document(dataDir + "input.docx", lo);
        //Save the document into DOCX
        dataDir = dataDir + "output.docx";
        doc.save(dataDir, SaveFormat.DOCX);
        // ExEnd:LoadOptionsUpdateDirtyFields
        System.out.println("\nUpdate the fields with the dirty attribute successfully.\nFile saved at " + dataDir);
    }

    public static void loadAndSaveEncryptedODT(String dataDir) throws Exception
    {
        // ExStart:LoadAndSaveEncryptedODT
        Document doc = new Document(dataDir + "encrypted.odt", new com.aspose.words.LoadOptions("password"));
        doc.save(dataDir + "out.odt", new OdtSaveOptions("newpassword"));
        // ExEnd:LoadAndSaveEncryptedODT
        System.out.println("\nLoad and save encrypted document successfully.\nFile saved at " + dataDir);
    }

    public static void verifyODTdocument(String dataDir) throws Exception
    {
        // ExStart:VerifyODTdocument
        FileFormatInfo info = FileFormatUtil.detectFileFormat(dataDir + "encrypted.odt");
        System.out.println(info.isEncrypted());
        // ExEnd:VerifyODTdocument
    }

    public static void convertShapeToOfficeMath(String dataDir) throws Exception
    {
        // ExStart:ConvertShapeToOfficeMath
        LoadOptions lo = new LoadOptions();
        lo.setConvertShapeToOfficeMath(true);

        // Specify load option to use previous default behaviour i.e. convert math shapes to office math ojects on loading stage.
        Document doc = new Document(dataDir + "OfficeMath.docx", lo);
        //Save the document into DOCX
        doc.save(dataDir + "ConvertShapeToOfficeMath_out.docx", SaveFormat.DOCX);
        // ExEnd:ConvertShapeToOfficeMath
    }

    public static void annotationsAtBlockLevel(String dataDir) throws Exception
    {
        // ExStart:AnnotationsAtBlockLevel
        LoadOptions options = new LoadOptions();
        options.setAnnotationsAtBlockLevel(false);
        Document doc = new Document(dataDir + "AnnotationsAtBlockLevel.docx", options);
        DocumentBuilder builder = new DocumentBuilder(doc);

        StructuredDocumentTag sdt = (StructuredDocumentTag)doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true).get(0);

        BookmarkStart start = builder.startBookmark("bm");
        BookmarkEnd end = builder.endBookmark("bm");

        sdt.getParentNode().insertBefore(start, sdt);
        sdt.getParentNode().insertAfter(end, sdt);

        //Save the document into DOCX
        doc.save(dataDir + "AnnotationsAtBlockLevel_out.docx", SaveFormat.DOCX);
        // ExEnd:AnnotationsAtBlockLevel
    }
}
