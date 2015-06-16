<?php
/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

use com\aspose\words\DocumentVisitor as JDocumentVisitor;

class FieldsHelper extends JDocumentVisitor {

    /**
     * Converts any fields of the specified type found in the descendants of the node into static text.
     *
     * @param compositeNode The node in which all descendants of the specified FieldType will be converted to static text.
     * @param targetFieldType The FieldType of the field to convert to static text.
     */

        public static function convertFieldsToStaticText($compositeNode, $targetFieldType) {

            $saveFormat = Java("com.aspose.words.SaveFormat");
            $originalNodeText = $compositeNode->toString($saveFormat->TEXT); //ExSkip

            $helper = new Java("com.aspose.words.FieldsHelper", $targetFieldType); // FieldsHelper(targetFieldType);
            $compositeNode->accept(helper);

            assert ($originalNodeText->equals($compositeNode->toString($saveFormat->TEXT))): // "Error: Text of the node converted differs from the original"; //ExSkip

            for (Node node : (Iterable<Node>)compositeNode.getChildNodes(NodeType.ANY, true)) //ExSkip
                assert (!(node instanceof FieldChar && ((FieldChar)node).getFieldType() == targetFieldType)) : "Error: A field node that should be removed still remains."; //ExSkip


        }


        private FieldsHelper(int targetFieldType)
        {
            mTargetFieldType = targetFieldType;
        }

        public int visitFieldStart(FieldStart fieldStart)
        {
            // We must keep track of the starts and ends of fields incase of any nested fields.
            if (fieldStart.getFieldType() == mTargetFieldType)
            {
                mFieldDepth++;
                fieldStart.remove();
            }
            else
            {
                // This removes the field start if it's inside a field that is being converted.
                CheckDepthAndRemoveNode(fieldStart);
            }

            return VisitorAction.CONTINUE;
        }

        public int visitFieldSeparator(FieldSeparator fieldSeparator)
        {
            // When visiting a field separator we should decrease the depth level.
            if (fieldSeparator.getFieldType() == mTargetFieldType)
            {
                mFieldDepth--;
                fieldSeparator.remove();
            }
            else
            {
                // This removes the field separator if it's inside a field that is being converted.
                CheckDepthAndRemoveNode(fieldSeparator);
            }

            return VisitorAction.CONTINUE;
        }

        public int visitFieldEnd(FieldEnd fieldEnd)
        {
            if (fieldEnd.getFieldType() == mTargetFieldType)
                fieldEnd.remove();
            else
                CheckDepthAndRemoveNode(fieldEnd); // This removes the field end if it's inside a field that is being converted.

            return VisitorAction.CONTINUE;
        }

        public int visitRun(Run run)
        {
            // Remove the run if it is between the FieldStart and FieldSeparator of the field being converted.
            CheckDepthAndRemoveNode(run);

            return VisitorAction.CONTINUE;
        }

        public int visitParagraphEnd(Paragraph paragraph)
        {
            if (mFieldDepth > 0)
            {
                // The field code that is being converted continues onto another paragraph. We
                // need to copy the remaining content from this paragraph onto the next paragraph.
                Node nextParagraph = paragraph.getNextSibling();

                // Skip ahead to the next available paragraph.
                while (nextParagraph != null && nextParagraph.getNodeType() != NodeType.PARAGRAPH)
                    nextParagraph = nextParagraph.getNextSibling();

                // Copy all of the nodes over. Keep a list of these nodes so we know not to remove them.
                while (paragraph.hasChildNodes())
                {
                    mNodesToSkip.add(paragraph.getLastChild());
                    ((Paragraph)nextParagraph).prependChild(paragraph.getLastChild());
                }

                paragraph.remove();
            }

            return VisitorAction.CONTINUE;
        }

        public int visitTableStart(Table table)
        {
            CheckDepthAndRemoveNode(table);

            return VisitorAction.CONTINUE;
        }

        /**
         * Checks whether the node is inside a field or should be skipped and then removes it if necessary.
         */
        private void CheckDepthAndRemoveNode(Node node)
        {
            if (mFieldDepth > 0 && !mNodesToSkip.contains(node))
                node.remove();
        }

        private int mFieldDepth = 0;
        private ArrayList mNodesToSkip = new ArrayList();
        private int mTargetFieldType;

}

class ReplaceFieldsWithStaticText  {

    public static function main() {

        // The path to the documents directory.
        //$dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/programmingwithdocuments/workingwithfields/replacefieldwithstatictext/data/";
        // Call a method to show how to convert all IF fields in a document to static text.
        //ReplaceFieldsWithStaticText::convertFieldsInDocument($dataDir);
        // Reload the document and this time convert all PAGE fields only encountered in the first body of the document.
        //ReplaceFieldsWithStaticText::convertFieldsInBody($dataDir);
        // Reload the document again and convert only the IF field in the last paragraph to static text.
        //ReplaceFieldsWithStaticText::convertFieldsInParagraph($dataDir);

        $obj = new FieldsHelper();
        echo '<pre>';
        echo java_inspect($obj);
        exit;
        echo $obj->hello();
        exit;

    }

    public static function convertFieldsInDocument($dataDir) {

        //ExStart:
        //ExId:FieldsToStaticTextDocument
        //ExSummary:Shows how to convert all fields of a specified type in a document to static text.
        $doc = new Java ("com.aspose.words.Document" , $dataDir . "TestFile.doc");

        // Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to static text.
        $fieldType = Java("com.aspose.words.FieldType");
        FieldsHelper::convertFieldsToStaticText($doc, $fieldType->FIELD_IF);

        // Save the document with fields transformed to disk.
        $doc->save($dataDir . "TestFileDocument Out.doc");
        //ExEnd

    }





}