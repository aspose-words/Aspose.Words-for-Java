package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.NodeType;
import com.aspose.words.OlePackage;
import com.aspose.words.Shape;
import com.aspose.words.examples.Utils;

import java.io.ByteArrayInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class DocumentBuilderInsertElements {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderInsertElements.class);

        insertOleObjectwithOlePackage(dataDir);
        GetAccessToOLEObjectRawData(dataDir);
    }

    public static void insertOleObjectwithOlePackage(String dataDir) throws Exception {
        // ExStart:InsertOleObjectwithOlePackage
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Path path = Paths.get(dataDir, "input.zip");
        byte[] bs = Files.readAllBytes(path);

        ByteArrayInputStream stream = new ByteArrayInputStream(bs);

        Shape shape = builder.insertOleObject(stream, "Package", true, null);
        OlePackage olePackage = shape.getOleFormat().getOlePackage();
        olePackage.setFileName("filename.zip");
        olePackage.setDisplayName("displayname.zip");
        dataDir = dataDir + "DocumentBuilderInsertOleObjectOlePackage_out.doc";
        doc.save(dataDir);

        // ExEnd:InsertOleObjectwithOlePackage
        System.out.println("\nOleObject using DocumentBuilder inserted successfully into a document.\nFile saved at " + dataDir);
    }
    
    public static void GetAccessToOLEObjectRawData(String dataDir) throws Exception
    {
        // ExStart:GetAccessToOLEObjectRawData
        // Load document with OLE object.
        Document doc = new Document(dataDir + "DocumentBuilderInsertTextInputFormField_out.doc");
        
        Shape oleShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);
        byte[] oleRawData = oleShape.getOleFormat().getRawData();
        // ExEnd:GetAccessToOLEObjectRawData
    }
}
