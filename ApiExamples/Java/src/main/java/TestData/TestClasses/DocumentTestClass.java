package TestData.TestClasses;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.Document;

import java.io.FileInputStream;

public class DocumentTestClass {
    private Document mDocument;
    private FileInputStream mDocumentStream;
    private byte[] mDocumentBytes;
    private String mDocumentUri;

    public DocumentTestClass() {
    }

    public DocumentTestClass(final Document doc, final FileInputStream docStream, final byte[] docBytes, final String docUri) {
        setDocument(doc);
        setDocumentStream(docStream);
        setDocumentBytes(docBytes);
        setDocumentUri(docUri);
    }

    public void setDocument(final Document value) {
        mDocument = value;
    }

    public void setDocumentStream(final FileInputStream value) {
        mDocumentStream = value;
    }

    public void setDocumentBytes(final byte[] value) {
        mDocumentBytes = value;
    }

    public void setDocumentUri(final String value) {
        mDocumentUri = value;
    }

    public Document getDocument() {
        return mDocument;
    }

    public FileInputStream getDocumentStream() {
        return mDocumentStream;
    }

    public byte[] getDocumentBytes() {
        return mDocumentBytes;
    }

    public String getDocumentUri() {
        return mDocumentUri;
    }
}
