package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.DocumentBuilder;
import com.aspose.words.FieldIf;

/**
 * Created by awaishafeez on 12/19/2017.
 */
public class EvaluateIFCondition {
    public static void main(String[] args) throws Exception {
        // ExStart:EvaluateIFCondition
        DocumentBuilder builder = new DocumentBuilder();
        FieldIf field = (FieldIf) builder.insertField("IF 1 = 1", null);

        int actualResult = field.evaluateCondition();
        switch (actualResult) {
            case 0:
                System.out.println("ERROR");
                break;
            case 1:
                System.out.println("TRUE");
                break;
            case 2:
                System.out.println("FALSE");
                break;
        }
        // ExEnd:EvaluateIFCondition
    }
}
