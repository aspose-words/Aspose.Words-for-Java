package TestData.TestClasses;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
    private String mDocumentString;

    public DocumentTestClass() {
    }

    public DocumentTestClass(final Document doc, final FileInputStream docStream, final byte[] docBytes, final String docString) {
        setDocument(doc);
        setDocumentStream(docStream);
        setDocumentBytes(docBytes);
        setDocumentString(docString);
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

    public void setDocumentString(final String value) {
        mDocumentString = value;
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

    public String getDocumentString() {
        return mDocumentString;
    }
}
