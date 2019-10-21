package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.OdtSaveMeasureUnit;
import com.aspose.words.OdtSaveOptions;
import com.aspose.words.OoxmlSaveOptions;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 9/18/2017.
 */
public class WorkingWithSaveOptions {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(WorkingWithSaveOptions.class);
        UpdateLastSavedTimeProperty(dataDir);
        SetMeasureUnitForODT(dataDir);

    }

    public static void UpdateLastSavedTimeProperty(String dataDir) throws Exception {
        // ExStart:UpdateLastSavedTimeProperty
        Document doc = new Document(dataDir + "Document.doc");

        OoxmlSaveOptions options = new OoxmlSaveOptions();
        options.setUpdateLastSavedTimeProperty(true);

        dataDir = dataDir + "UpdateLastSavedTimeProperty_out.docx";

        // Save the document to disk.
        doc.save(dataDir, options);
        // ExEnd:UpdateLastSavedTimeProperty
        System.out.println("\nUpdated Last Saved Time Property successfully.");
    }

    public static void SetMeasureUnitForODT(String dataDir) throws Exception {
        // ExStart:SetMeasureUnitForODT
        //Load the Word document
        Document doc = new Document(dataDir + "Document.doc");

        //Open Office uses centimeters when specifying lengths, widths and other measurable formatting
        //and content properties in documents whereas MS Office uses inches.

        OdtSaveOptions saveOptions = new OdtSaveOptions();
        saveOptions.setMeasureUnit(OdtSaveMeasureUnit.INCHES);

        //Save the document into ODT
        doc.save(dataDir + "MeasureUnit_out.odt", saveOptions);
        // ExEnd:SetMeasureUnitForODT
        System.out.println("\nSet MeasureUnit for ODT successfully.\nFile saved at " + dataDir);
    }
}


