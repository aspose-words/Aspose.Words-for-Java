package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Section;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.loading_saving.*;
import java.io.*;

public class SplitDocument {

	public static void main(String[] args) throws Exception{
		// TODO Auto-generated method stub

		// The path to the documents directory.
		String dataDir = Utils.getDataDir(SplitDocument.class);
		
		SplitDocumentBySections(dataDir);
		SplitDocumentPageByPage(dataDir);
		SplitDocumentByPageRange(dataDir);
		MergeDocuments(dataDir);
	}

	public static void SplitDocumentBySections (String dataDir) throws Exception {
		// ExStart:SplitDocumentBySections
		// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-Java
		// Open a Word document
		Document doc = new Document(dataDir + "TestFile (Split).docx");

		for (int i = 0; i < doc.getSections().getCount(); i++)
		{
		    // Split a document into smaller parts, in this instance split by section
		    Section section = doc.getSections().get(i).deepClone();

		    Document newDoc = new Document();
		    newDoc.getSections().clear();

		    Section newSection = (Section) newDoc.importNode(section, true);
		    newDoc.getSections().add(newSection);

		    // Save each section as a separate document
		    newDoc.save(dataDir + "SplitDocumentBySectionsOut_" + i + ".docx");
		}
		// ExEnd:SplitDocumentBySections
	}
	
	public static void SplitDocumentPageByPage (String dataDir) throws Exception {
		// ExStart:SplitDocumentPageByPage
		// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-Java
		// Open a Word document
		Document doc = new Document(dataDir + "TestFile (Split).docx");

		// Split nodes in the document into separate pages
		DocumentPageSplitter splitter = new DocumentPageSplitter(doc);

		// Save each page as a separate document
		for (int page = 1; page <= doc.getPageCount(); page++)
		{
		    Document pageDoc = splitter.getDocumentOfPage(page);
		    pageDoc.save(dataDir + "SplitDocumentPageByPageOut_" + page + ".docx");
		}
		// ExEnd:SplitDocumentPageByPage
	}
	
	public static void SplitDocumentByPageRange (String dataDir) throws Exception {
		// ExStart:SplitDocumentByPageRange
		// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-Java
		// Open a Word document
		Document doc = new Document(dataDir + "TestFile (Split).docx");

		// Split nodes in the document into separate pages
		DocumentPageSplitter splitter = new DocumentPageSplitter(doc);
		 
		// Get part of the document
		Document pageDoc = splitter.getDocumentOfPageRange(3,6);
		pageDoc.save(dataDir + "SplitDocumentByPageRangeOut.docx");
		// ExEnd:SplitDocumentByPageRange
	}
	
	//ExStart: MergeDocuments
	public static void MergeDocuments(String dataDir) throws Exception{
	    // Find documents using for merge
		File f = new File(dataDir);

		FilenameFilter filter = new FilenameFilter() {
		        @Override
		        public boolean accept(File f, String name) {
		            return name.endsWith(".docx");
		        }
		    };

		String[] documentPaths = f.list(filter);
		
	    String sourceDocumentPath = dataDir + documentPaths[0];

	    // Open the first part of the resulting document
	    Document sourceDoc = new Document(sourceDocumentPath);

	    // Create a new resulting document
	    Document mergedDoc = new Document();
	    DocumentBuilder mergedDocBuilder = new DocumentBuilder(mergedDoc);

	    // Merge document parts one by one
	    for (String documentPath : documentPaths)
	    {
	    	String documentPathFull = dataDir + documentPath;
	        if (documentPathFull == sourceDocumentPath)
	            continue;

	        mergedDocBuilder.moveToDocumentEnd();
	        mergedDocBuilder.insertDocument(sourceDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
	        sourceDoc = new Document(documentPathFull);
	    }

	    // Save the output file
	    mergedDoc.save(dataDir + "MergeDocuments_out.docx");
	}
	//ExEnd: MergeDocuments
}
