package com.aspose.words.examples.loading_saving;

import java.io.FileInputStream;
import java.io.InputStream;

import com.aspose.words.Document;
import com.aspose.words.LoadOptions;
import com.aspose.words.examples.Utils;

public class OpenDocument {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(OpenDocument.class);

		OpenFromFile(dataDir);
		OpenEncryptedDocument(dataDir);
		OpenFromStream(dataDir);
	}

	public static void OpenFromFile(String dataDir) throws Exception {
		// ExStart:OpenFromFile
		// For complete examples and data files, please go to
		// https://github.com/aspose-words/Aspose.Words-for-Java
		String fileName = "Document.docx";

		// Load the document from the absolute path on disk.
		Document doc = new Document(dataDir + fileName);
		// ExEnd:OpenFromFile
		System.out.println("Document loaded successfully.");
	}

	public static void OpenEncryptedDocument(String dataDir) throws Exception {
		// ExStart: OpenEncryptedDocument
		// For complete examples and data files, please go to
		// https://github.com/aspose-words/Aspose.Words-for-Java
		// Load the encrypted document from the absolute path on disk.
		Document doc = new Document(dataDir + "LoadEncrypted.docx", new LoadOptions("aspose"));
		// ExEnd: OpenEncryptedDocument
		System.out.println("Encrypted document loaded successfully.");
	}

	public static void OpenFromStream(String dataDir) throws Exception {
		// ExStart: OpenFromStream
		// For complete examples and data files, please go to
		// https://github.com/aspose-words/Aspose.Words-for-Java
		String filename = "Document.docx";

		// Open the stream. Read only access is enough for Aspose.Words to load a
		// document.
		InputStream in = new FileInputStream(dataDir + filename);

		// Load the entire document into memory.
		Document doc = new Document(in);
		System.out.println("Document opened. Total pages are " + doc.getPageCount());
		// You can close the stream now, it is no longer needed because the document is
		// in memory.
		in.close();
		// ExEnd: OpenFromStream
	}

}
