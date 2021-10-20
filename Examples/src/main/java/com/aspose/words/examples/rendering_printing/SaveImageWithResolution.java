package com.aspose.words.examples.rendering_printing;

import com.aspose.words.ImageSaveOptions;
import com.aspose.words.PageSet;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 5/29/2017.
 */
public class SaveImageWithResolution {
    public static void main(String[] args) throws Exception {
        //ExStart:SetHorizontalAndVerticalImageResolution
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SaveImageWithResolution.class);
        com.aspose.words.Document doc = new com.aspose.words.Document(dataDir + "TestFile.doc");

        //Renders a page of a Word document into a PNG image at a specific horizontal and vertical resolution.
        ImageSaveOptions options = new ImageSaveOptions(com.aspose.words.SaveFormat.PNG);
        options.setHorizontalResolution(300);
        options.setVerticalResolution(300);
        options.setPageSet(new PageSet(0, 1));

        doc.save(dataDir + "Rendering.SaveToImageResolution Out.png", options);
        //ExEnd:SetHorizontalAndVerticalImageResolution
    }
}
