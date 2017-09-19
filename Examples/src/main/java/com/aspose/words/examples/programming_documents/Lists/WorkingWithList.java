package com.aspose.words.examples.programming_documents.Lists;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 9/18/2017.
 */
public class WorkingWithList {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(WorkingWithList.class);
        SetRestartAtEachSection(dataDir);
    }

    public static void SetRestartAtEachSection(String dataDir) throws Exception {
        // ExStart:SetRestartAtEachSection
        Document doc = new Document();

        doc.getLists().add(ListTemplate.NUMBER_DEFAULT);

        com.aspose.words.List list = doc.getLists().get(0);

        // Set true to specify that the list has to be restarted at each section.
        list.isRestartAtEachSection(true);

        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getListFormat().setList(list);

        for (int i = 1; i < 45; i++) {
            builder.writeln(String.format("List Item " + i));

            // Insert section break.
            if (i == 15)
                builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        }
        builder.getListFormat().removeNumbers();
        // IsRestartAtEachSection will be written only if compliance is higher then OoxmlComplianceCore.Ecma376
        OoxmlSaveOptions options = new OoxmlSaveOptions();
        options.setCompliance(OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);

        dataDir = dataDir + "RestartAtEachSection_out.docx";

        // Save the document to disk.
        doc.save(dataDir, options);
        // ExEnd:SetRestartAtEachSection
        System.out.println("\nDocument is saved successfully.");
    }
}
