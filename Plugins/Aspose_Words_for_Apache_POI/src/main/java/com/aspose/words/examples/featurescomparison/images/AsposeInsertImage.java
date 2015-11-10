package com.aspose.words.examples.featurescomparison.images;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.WrapType;
import com.aspose.words.examples.Utils;

public class AsposeInsertImage
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeInsertImage.class);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertImage(dataDir + "background.jpg");
        builder.insertImage(dataDir + "background.jpg",
                RelativeHorizontalPosition.MARGIN,
                100,
                RelativeVerticalPosition.MARGIN,
                200,
                200,
                100,
                WrapType.SQUARE);

        doc.save(dataDir + "Aspose_InsertImage.docx");

        System.out.println("Process Completed Successfully");
    }
}
