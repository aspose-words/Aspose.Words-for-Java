package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.DocumentBuilder;
import com.aspose.words.Field;

public class SpecifylocaleAtFieldlevel {
    public static void main(String[] args) throws Exception {
        // ExStart:SpecifylocaleAtFieldlevel
        DocumentBuilder builder = new DocumentBuilder();
        Field field = builder.insertField("=1", null);
        field.setLocaleId(1027);
        // ExEnd:SpecifylocaleAtFieldlevel
    }
}