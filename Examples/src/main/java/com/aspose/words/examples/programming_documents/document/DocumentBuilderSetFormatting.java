package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.examples.Utils;
import com.aspose.words.examples.loading_saving.WorkingWithTxt;
import com.aspose.words.ParagraphFormat;
import com.aspose.words.Document;

public class DocumentBuilderSetFormatting {
    public static void main(String[] args) throws Exception {
        String dataDir = Utils.getDataDir(DocumentBuilderSetFormatting.class);

        setAsianTypographyLinebreakGroupProp(dataDir);
    }

    public static void setAsianTypographyLinebreakGroupProp(String dataDir) throws Exception
    {
        // ExStart:SetAsianTypographyLinebreakGroupProp
        Document doc = new Document(dataDir + "Input.docx");

        ParagraphFormat format = doc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat();
        format.setFarEastLineBreakControl(false);
        format.setWordWrap(true);
        format.setHangingPunctuation(false);

        dataDir = dataDir + "SetAsianTypographyLinebreakGroupProp_out.doc";
        doc.save(dataDir);
        // ExEnd:SetAsianTypographyLinebreakGroupProp
        System.out.println("\nParagraphFormat properties for Asian Typography line break group are set successfully.\nFile saved at " + dataDir);
    }
}
