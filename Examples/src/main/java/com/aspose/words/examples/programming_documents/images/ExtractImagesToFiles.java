package com.aspose.words.examples.programming_documents.images;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class ExtractImagesToFiles {
    public static void main(String[] args) throws Exception {
        
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ExtractImagesToFiles.class);

        Document doc = new Document(dataDir + "Image.SampleImages.doc");

        NodeCollection<Shape> shapes = (NodeCollection<Shape>) doc.getChildNodes(NodeType.SHAPE, true);
        int imageIndex = 0;
        for (Shape shape : shapes
                ) {
            if (shape.hasImage()) {
                String imageFileName = String.format(
                        "Image.ExportImages.{0}_out_{1}", imageIndex, FileFormatUtil.imageTypeToExtension(shape.getImageData().getImageType()));
                shape.getImageData().save(dataDir + imageFileName);
                imageIndex++;
            }
        }
    }
}