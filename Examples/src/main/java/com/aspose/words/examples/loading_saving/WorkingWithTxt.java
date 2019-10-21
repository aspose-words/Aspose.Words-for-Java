package com.aspose.words.examples.loading_saving;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class WorkingWithTxt {
    public static void main(String[] args) throws Exception {
        String dataDir = Utils.getDataDir(WorkingWithTxt.class);

        saveAsTxt(dataDir);
        addBidiMarks(dataDir);
        detectNumberingWithWhitespaces(dataDir);
        handleSpacesOptions(dataDir);
        exportHeadersFootersMode(dataDir);
        useTabCharacterPerLevelForListIndentation(dataDir);
        useSpaceCharacterPerLevelForListIndentation(dataDir);
    }

    public static void saveAsTxt(String dataDir) throws Exception {
        // ExStart:SaveAsTxt
        Document doc = new Document(dataDir + "Document.doc");
        dataDir = dataDir + "Document.ConvertToTxt_out.txt";
        doc.save(dataDir);
        // ExEnd:SaveAsTxt
        System.out.println("\nDocument saved as TXT.\nFile saved at " + dataDir);
    }

    public static void addBidiMarks(String dataDir) throws Exception {
        // ExStart:AddBidiMarks
        Document doc = new Document(dataDir + "Input.docx");
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        //The default value is false.
        saveOptions.setAddBidiMarks(true);

        dataDir = dataDir + "Document.AddBidiMarks_out.txt";
        doc.save(dataDir, saveOptions);
        // ExEnd:AddBidiMarks
        System.out.println("\nAdd bi-directional marks set successfully.\nFile saved at " + dataDir);
    }

    public static void detectNumberingWithWhitespaces(String dataDir) throws Exception {
        //ExStart:DetectNumberingWithWhitespaces
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.setDetectNumberingWithWhitespaces(false);

        Document doc = new Document(dataDir + "LoadTxt.txt", loadOptions);

        dataDir = dataDir + "DetectNumberingWithWhitespaces_out.docx";
        doc.save(dataDir);
        //ExEnd:DetectNumberingWithWhitespaces
        System.out.println("\nDetect number with whitespaces successfully.\nFile saved at " + dataDir);
    }

    public static void handleSpacesOptions(String dataDir) throws Exception {
        //ExStart:HandleSpacesOptions
        TxtLoadOptions loadOptions = new TxtLoadOptions();

        loadOptions.setLeadingSpacesOptions(TxtLeadingSpacesOptions.TRIM);
        loadOptions.setTrailingSpacesOptions(TxtTrailingSpacesOptions.TRIM);
        Document doc = new Document(dataDir + "LoadTxt.txt", loadOptions);

        dataDir = dataDir + "HandleSpacesOptions_out.docx";
        doc.save(dataDir);
        //ExEnd:HandleSpacesOptions
        System.out.println("\nTrim leading and trailing spaces while importing text document.\nFile saved at " + dataDir);
    }

    public static void exportHeadersFootersMode(String dataDir) throws Exception {
        //ExStart:ExportHeadersFootersMode

        Document doc = new Document(dataDir + "TxtExportHeadersFootersMode.docx");

        TxtSaveOptions options = new TxtSaveOptions();
        options.setSaveFormat(SaveFormat.TEXT);

        // All headers and footers are placed at the very end of the output document.
        options.setExportHeadersFootersMode(TxtExportHeadersFootersMode.ALL_AT_END);
        doc.save(dataDir + "outputFileNameA.txt", options);

        // Only primary headers and footers are exported at the beginning and end of each section.
        options.setExportHeadersFootersMode(TxtExportHeadersFootersMode.PRIMARY_ONLY);
        doc.save(dataDir + "outputFileNameB.txt", options);

        // No headers and footers are exported.
        options.setExportHeadersFootersMode(TxtExportHeadersFootersMode.NONE);
        doc.save(dataDir + "outputFileNameC.txt", options);

        //ExEnd:ExportHeadersFootersMode
        System.out.println("\nExport text files with TxtExportHeadersFootersMode.\nFiles saved at " + dataDir);
    }

    public static void useTabCharacterPerLevelForListIndentation(String dataDir) throws Exception {
        //ExStart:useTabCharacterPerLevelForListIndentation
        Document doc = new Document(dataDir + "Input.docx");

        TxtSaveOptions options = new TxtSaveOptions();
        options.getListIndentation().setCount(1);
        options.getListIndentation().setCharacter('\t');

        doc.save(dataDir + "output.txt", options);
        //ExEnd:useTabCharacterPerLevelForListIndentation
    }

    public static void useSpaceCharacterPerLevelForListIndentation(String dataDir) throws Exception {
        //ExStart:useSpaceCharacterPerLevelForListIndentation
        Document doc = new Document(dataDir + "Input.docx");

        TxtSaveOptions options = new TxtSaveOptions();
        options.getListIndentation().setCount(3);
        options.getListIndentation().setCharacter(' ');

        doc.save(dataDir + "output.txt", options);
        //ExEnd:useSpaceCharacterPerLevelForListIndentation
    }

    public static void defaultLevelForListIndentation(String dataDir) throws Exception {
        //ExStart:defaultLevelForListIndentation
        Document doc = new Document(dataDir + "Input.docx");
        doc.save(dataDir + "output1.txt");

        Document doc2 = new Document("Input.docx");
        TxtSaveOptions options = new TxtSaveOptions();
        doc2.save(dataDir + "output2.txt", options);
        //ExEnd:defaultLevelForListIndentation
    }
    
    public static void DocumentTextDirection(String dataDir) throws Exception{
    	//ExStart: DocumentTextDirection
    	TxtLoadOptions loadOptions = new TxtLoadOptions();
    	loadOptions.setDocumentDirection(DocumentDirection.AUTO);

    	Document doc = new Document(dataDir + "arabic.txt", loadOptions);

    	Paragraph paragraph = doc.getFirstSection().getBody().getFirstParagraph();
    	System.out.println(paragraph.getParagraphFormat().getBidi());

    	dataDir = dataDir + "DocumentDirection_out.docx";
    	doc.save(dataDir);
    	//ExEnd: DocumentTextDirection
    }
}
