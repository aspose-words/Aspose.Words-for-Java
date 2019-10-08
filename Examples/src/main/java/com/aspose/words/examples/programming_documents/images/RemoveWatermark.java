package com.aspose.words.examples.programming_documents.images;

import com.aspose.words.Document;
import com.aspose.words.HeaderFooter;
import com.aspose.words.NodeType;
import com.aspose.words.Shape;
import com.aspose.words.examples.Utils;

public class RemoveWatermark {

    // ExStart:RemoveWatermark
    private static final String dataDir = Utils.getDataDir(RemoveWatermark.class);

    public static void main(String[] args) throws Exception {
        Document doc = new Document(dataDir + "RemoveWatermark.docx");
        removeWatermarkText(doc);
        doc.save(dataDir + "RemoveWatermark_out.doc");
    }

    private static void removeWatermarkText(Document doc) throws Exception {
        for (HeaderFooter hf : (Iterable<HeaderFooter>) doc.getChildNodes(NodeType.HEADER_FOOTER, true)) {
            for (Shape shape : (Iterable<Shape>) hf.getChildNodes(NodeType.SHAPE, true)) {
                if (shape.getName().contains("WaterMark"))
                    shape.remove();
            }
        }
    }
    // ExEnd:RemoveWatermark
}