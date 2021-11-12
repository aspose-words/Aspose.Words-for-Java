package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.FileFormatInfo;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.LoadOptions;
import com.aspose.words.MsWordVersion;
import com.aspose.words.OdtSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class SpecifyLoadOptions {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(SpecifyLoadOptions.class);

		setMSWordVersion(dataDir);
		loadOptionsUpdateDirtyFields(dataDir);
		loadAndSaveEncryptedODT(dataDir);
		verifyODTdocument(dataDir);
		convertShapeToOfficeMath(dataDir);
		SetTempFolder(dataDir);
		LoadOptionsEncoding(dataDir);

	}

	public static void loadOptionsUpdateDirtyFields(String dataDir) throws Exception {
		// ExStart:LoadOptionsUpdateDirtyFields
		LoadOptions lo = new LoadOptions();
		// Update the fields with the dirty attribute
		lo.setUpdateDirtyFields(true);
		// Load the Word document
		Document doc = new Document(dataDir + "input.docx", lo);
		// Save the document into DOCX
		dataDir = dataDir + "output.docx";
		doc.save(dataDir, SaveFormat.DOCX);
		// ExEnd:LoadOptionsUpdateDirtyFields
		System.out.println("\nUpdate the fields with the dirty attribute successfully.\nFile saved at " + dataDir);
	}

	public static void loadAndSaveEncryptedODT(String dataDir) throws Exception {
		// ExStart:LoadAndSaveEncryptedODT
		Document doc = new Document(dataDir + "encrypted.odt", new com.aspose.words.LoadOptions("password"));
		doc.save(dataDir + "out.odt", new OdtSaveOptions("newpassword"));
		// ExEnd:LoadAndSaveEncryptedODT
		System.out.println("\nLoad and save encrypted document successfully.\nFile saved at " + dataDir);
	}

	public static void verifyODTdocument(String dataDir) throws Exception {
		// ExStart:VerifyODTdocument
		FileFormatInfo info = FileFormatUtil.detectFileFormat(dataDir + "encrypted.odt");
		System.out.println(info.isEncrypted());
		// ExEnd:VerifyODTdocument
	}

	public static void convertShapeToOfficeMath(String dataDir) throws Exception {
		// ExStart:ConvertShapeToOfficeMath
		LoadOptions lo = new LoadOptions();
		lo.setConvertShapeToOfficeMath(true);

		// Specify load option to use previous default behaviour i.e. convert math
		// shapes to office math ojects on loading stage.
		Document doc = new Document(dataDir + "OfficeMath.docx", lo);
		// Save the document into DOCX
		doc.save(dataDir + "ConvertShapeToOfficeMath_out.docx", SaveFormat.DOCX);
		// ExEnd:ConvertShapeToOfficeMath
	}

	public static void setMSWordVersion(String dataDir) throws Exception {
		// ExStart:SetMSWordVersion
		// Specify load option to specify MS Word version
		LoadOptions loadOptions = new LoadOptions();
		loadOptions.setMswVersion(MsWordVersion.WORD_2003);

		Document doc = new Document(dataDir + "document.doc", loadOptions);
		doc.save(dataDir + "Word2003_out.docx");
		// ExEnd:SetMSWordVersion
	}

	public static void SetTempFolder(String dataDir) throws Exception {
		// ExStart:SetTempFolder
		// Specify LoadOptions to set Temp Folder
		LoadOptions lo = new LoadOptions();
		lo.setTempFolder("C:\\TempFolder\\");

		Document doc = new Document(dataDir + "document.doc", lo);
		// ExEnd:SetTempFolder
	}

	public static void LoadOptionsEncoding(String dataDir) throws Exception {
		// ExStart:LoadOptionsEncoding
		// Set the Encoding attribute in a LoadOptions object to override the
		// automatically chosen encoding with the one we know to be correct
		LoadOptions loadOptions = new LoadOptions();
		loadOptions.setEncoding(java.nio.charset.Charset.forName("UTF-8"));

		Document doc = new Document(dataDir + "Encoded in UTF-8.txt", loadOptions);
		// ExEnd:LoadOptionsEncoding
	}
}
