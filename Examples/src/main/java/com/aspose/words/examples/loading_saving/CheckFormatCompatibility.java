package com.aspose.words.examples.loading_saving;

import com.aspose.words.FileFormatInfo;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.LoadFormat;
import com.aspose.words.examples.Utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class CheckFormatCompatibility {
	public static void main(String[] args) throws Exception {
		// ExStart:CheckFormatCompatibility
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(CheckFormatCompatibility.class);

		String supportedDir = dataDir + "OutSupported" + File.separator;
		String unknownDir = dataDir + "OutUnknown" + File.separator;
		String encryptedDir = dataDir + "OutEncrypted" + File.separator;
		String pre97Dir = dataDir + "OutPre97" + File.separator;

		//ExStart:GetListOfFilesInFolder
		File[] fileList = new File(dataDir).listFiles();
		//ExEnd:GetListOfFilesInFolder

		// Loop through all found files.
		for (File file : fileList) {
			if (file.isDirectory())
				continue;

			// Extract and display the file name without the path.
			String nameOnly = file.getName();
			System.out.print(nameOnly);

			// Check the file format and move the file to the appropriate folder.
			String fileName = file.getPath();
			FileFormatInfo info = FileFormatUtil.detectFileFormat(fileName);

			// Display the document type.
			switch (info.getLoadFormat()) {
			case LoadFormat.DOC:
				System.out.println("\tMicrosoft Word 97-2003 document.");
				break;
			case LoadFormat.DOT:
				System.out.println("\tMicrosoft Word 97-2003 template.");
				break;
			case LoadFormat.DOCX:
				System.out.println("\tOffice Open XML WordprocessingML Macro-Free Document.");
				break;
			case LoadFormat.DOCM:
				System.out.println("\tOffice Open XML WordprocessingML Macro-Enabled Document.");
				break;
			case LoadFormat.DOTX:
				System.out.println("\tOffice Open XML WordprocessingML Macro-Free Template.");
				break;
			case LoadFormat.DOTM:
				System.out.println("\tOffice Open XML WordprocessingML Macro-Enabled Template.");
				break;
			case LoadFormat.FLAT_OPC:
				System.out.println("\tFlat OPC document.");
				break;
			case LoadFormat.RTF:
				System.out.println("\tRTF format.");
				break;
			case LoadFormat.WORD_ML:
				System.out.println("\tMicrosoft Word 2003 WordprocessingML format.");
				break;
			case LoadFormat.HTML:
				System.out.println("\tHTML format.");
				break;
			case LoadFormat.MHTML:
				System.out.println("\tMHTML (Web archive) format.");
				break;
			case LoadFormat.ODT:
				System.out.println("\tOpenDocument Text.");
				break;
			case LoadFormat.OTT:
				System.out.println("\tOpenDocument Text Template.");
				break;
			case LoadFormat.DOC_PRE_WORD_60:
				System.out.println("\tMS Word 6 or Word 95 format.");
				break;
			case LoadFormat.UNKNOWN:
			default:
				System.out.println("\tUnknown format.");
				break;
			}

			// Now copy the document into the appropriate folder.
			if (info.isEncrypted()) {
				System.out.println("\tAn encrypted document.");
				fileCopy(fileName, new File(encryptedDir, nameOnly).getPath());
			} else {
				switch (info.getLoadFormat()) {
				case LoadFormat.DOC_PRE_WORD_60:
					fileCopy(fileName, new File(pre97Dir + nameOnly).getPath());
					break;
				case LoadFormat.UNKNOWN:
					fileCopy(fileName, new File(unknownDir + nameOnly).getPath());
					break;
				default:
					fileCopy(fileName, new File(supportedDir + nameOnly).getPath());
					break;
				}
			}
		}
	}// ExEnd:CheckFormatCompatibility

	// ExStart:fileCopy
	private static void fileCopy(String sourceFileName, String destinationFileName) throws Exception {
		File sourceFile = new File(sourceFileName);
		File destinationFile = new File(destinationFileName);

		File directoryFile = new File(destinationFile.getParent());
		if (!directoryFile.exists())
			directoryFile.mkdir();

		FileInputStream fis = null;
		FileOutputStream fos = null;

		try {
			fis = new FileInputStream(sourceFile);
			fos = new FileOutputStream(destinationFile);

			byte[] buffer = new byte[8192];
			int bytesRead;
			while ((bytesRead = fis.read(buffer)) != -1)
				fos.write(buffer, 0, bytesRead);
		} finally {
			if (fis != null)
				fis.close();
			if (fos != null)
				fos.close();
		}

	}
	// ExEnd:fileCopy

}
