package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.awt.*;


public class DocumentBuilderApplyBordersAndShadingToParagraph {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderApplyBordersAndShadingToParagraph.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set paragraph borders
        BorderCollection borders = builder.getParagraphFormat().getBorders();
        borders.setDistanceFromText(20);
        borders.getByBorderType(BorderType.LEFT).setLineStyle(LineStyle.DOUBLE);
        borders.getByBorderType(BorderType.RIGHT).setLineStyle(LineStyle.DOUBLE);
        borders.getByBorderType(BorderType.TOP).setLineStyle(LineStyle.DOUBLE);
        borders.getByBorderType(BorderType.BOTTOM).setLineStyle(LineStyle.DOUBLE);
        // Set paragraph shading
        Shading shading = builder.getParagraphFormat().getShading();
        shading.setTexture(TextureIndex.TEXTURE_DIAGONAL_CROSS);
        shading.setBackgroundPatternColor(Color.YELLOW);
        shading.setForegroundPatternColor(Color.GREEN);

        builder.write("I'm a formatted paragraph with double border and nice shading.");
        doc.save(dataDir + "output.doc");



    }
}
