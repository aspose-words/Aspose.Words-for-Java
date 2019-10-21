package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.Font;
import com.aspose.words.RunCollection;
import com.aspose.words.TextDmlEffect;
import com.aspose.words.examples.Utils;

public class CheckDMLTextEffect {

    public static void main(String[] args) throws Exception {
        // ExStart: CheckDMLTextEffect
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(CheckDMLTextEffect.class);

        // Initialize document.
        Document doc = new Document(dataDir + "Document.doc");
        RunCollection runs = doc.getFirstSection().getBody().getFirstParagraph().getRuns();

        Font runFont = runs.get(0).getFont();

        // One run might have several Dml text effects applied.
        System.out.println(runFont.hasDmlEffect(TextDmlEffect.SHADOW));
        System.out.println(runFont.hasDmlEffect(TextDmlEffect.EFFECT_3_D));
        System.out.println(runFont.hasDmlEffect(TextDmlEffect.REFLECTION));
        System.out.println(runFont.hasDmlEffect(TextDmlEffect.OUTLINE));
        System.out.println(runFont.hasDmlEffect(TextDmlEffect.FILL));
        // ExEnd: CheckDMLTextEffect
    }

}
