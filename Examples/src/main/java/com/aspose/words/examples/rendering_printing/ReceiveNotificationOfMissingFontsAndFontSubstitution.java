package com.aspose.words.examples.rendering_printing;

import com.aspose.words.Document;
import com.aspose.words.FontSettings;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.examples.Utils;

public class ReceiveNotificationOfMissingFontsAndFontSubstitution {

    private static final String dataDir = Utils.getSharedDataDir(ReceiveNotificationOfMissingFontsAndFontSubstitution.class) + "RenderingAndPrinting/";

    public static void main(String[] args) throws Exception {
        //ExStart:Main
        // Load the document to render.
        Document doc = new Document(dataDir + "Rendering.doc");

        // We can choose the default font to use in the case of any missing fonts.
        FontSettings.getDefaultInstance().getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");

        // For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
        // find any fonts in the specified directory, then during rendering the fonts in the document will be subsuited with the default
        // font specified under FontSettings.DefaultFontName. We can pick up on this subsuition using our callback.
        FontSettings.getDefaultInstance().setFontsFolder("", false);

        // Create a new class implementing IWarningCallback which collect any warnings produced during document save.
        HandleDocumentWarnings callback = new HandleDocumentWarnings();

        doc.setWarningCallback(callback);

        // Pass the save options along with the save path to the save method.
        doc.save(dataDir + "Rendering.MissingFontNotification Out.pdf");

        getNotificationBeforeSaving(doc);
        //ExEnd:Main
    }

    public static void getNotificationBeforeSaving(Document doc) throws Exception {
        //ExStart:GetNotificationBeforeSaving
        doc.updatePageLayout();

        // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class.
        HandleDocumentWarnings callback = new HandleDocumentWarnings();
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        doc.setWarningCallback(callback);

        // Even though the document was rendered previously, any save warnings are notified to the user during document save.
        doc.save(dataDir + "Rendering.FontsNotificationUpdatePageLayout Out.pdf");
        //ExEnd:GetNotificationBeforeSaving
    }
}