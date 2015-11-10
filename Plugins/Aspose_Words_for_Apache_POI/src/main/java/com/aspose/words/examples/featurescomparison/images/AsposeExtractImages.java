package com.aspose.words.examples.featurescomparison.images;

import com.aspose.words.Document;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Shape;
import com.aspose.words.examples.Utils;

public class AsposeExtractImages
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeExtractImages.class);

        Document doc = new Document(dataDir + "document.doc");

        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);
        int imageIndex = 0;
        for (Shape shape : (Iterable<Shape>) shapes)
        {
            if (shape.hasImage())
            {
                String imageFileName = java.text.MessageFormat.format(
                                "Aspose.Images.{0}{1}", imageIndex, FileFormatUtil
                                                .imageTypeToExtension(shape.getImageData()
                                                                .getImageType()));
                shape.getImageData().save(dataDir + imageFileName);

                imageIndex++;
            }
        }
    }
}
