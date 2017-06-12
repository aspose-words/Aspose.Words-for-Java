package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.HtmlLoadOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 6/12/2017.
 */
public class LoadAndSaveHtmlFormFieldasContentControlinDOCX {
    // The path to the documents directory.
    private static final String dataDir = Utils.getDataDir(LoadAndSaveHtmlFormFieldasContentControlinDOCX.class);

    public static void main(String[] args) throws Exception {

        HtmlLoadOptions lo = new HtmlLoadOptions();
        // lo.PreferredControlType = HtmlControlType.StructuredDocumentTag;

        //Load the HTML document
        Document doc = new Document(dataDir + "input.html", lo);

        //Save the HTML document as DOCX
        doc.save(dataDir + "output.docx", SaveFormat.DOCX);
        System.out.println("Html form fields are exported as content control successfully.");
    }

}