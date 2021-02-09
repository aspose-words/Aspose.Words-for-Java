package TestData.TestBuilders;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import TestData.TestClasses.DocumentTestClass;
import com.aspose.words.Document;

import java.io.FileInputStream;

public class DocumentTestBuilder {
    private Document mDocument;
    private FileInputStream mDocumentStream;
    private byte[] mDocumentBytes;
    private String mDocumentString;

    public DocumentTestBuilder() throws Exception {
        mDocument = new Document();
        mDocumentStream = null;
        mDocumentBytes = new byte[0];
        mDocumentString = "";
    }

    public DocumentTestBuilder withDocument(final Document doc) {
        mDocument = doc;
        return this;
    }

    public DocumentTestBuilder withDocumentStream(final FileInputStream stream) {
        mDocumentStream = stream;
        return this;
    }

    public DocumentTestBuilder withDocumentBytes(final byte[] docBytes) {
        mDocumentBytes = docBytes;
        return this;
    }

    public DocumentTestBuilder withDocumentString(final String docUri) {
        mDocumentString = docUri;
        return this;
    }

    public DocumentTestClass build() {
        return new DocumentTestClass(mDocument, mDocumentStream, mDocumentBytes, mDocumentString);
    }
}
