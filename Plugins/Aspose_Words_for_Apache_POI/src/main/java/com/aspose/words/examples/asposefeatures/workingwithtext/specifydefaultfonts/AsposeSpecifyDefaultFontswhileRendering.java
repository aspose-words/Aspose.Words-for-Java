package com.aspose.words.examples.asposefeatures.workingwithtext.specifydefaultfonts;

import com.aspose.words.Document;
import com.aspose.words.FontSettings;
import com.aspose.words.examples.Utils;

public class AsposeSpecifyDefaultFontswhileRendering
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeSpecifyDefaultFontswhileRendering.class);

        Document doc = new Document(dataDir + "document.doc");

        // If the default font defined here cannot be found during rendering then the closest font on the machine is used instead.
        FontSettings.setDefaultFontName("Arial Unicode MS");

        // Now the set default font is used in place of any missing fonts during any rendering calls.
        doc.save(dataDir + "AsposeSetDefaultFont_Out.pdf");
        doc.save(dataDir + "AsposeSetDefaultFont_Out.xps");

        System.out.println("Process Completed Successfully");
    }
}
