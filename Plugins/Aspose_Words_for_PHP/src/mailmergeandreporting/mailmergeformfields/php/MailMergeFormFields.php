<?php

/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

use com\aspose\words\Document as Document;
use com\aspose\words\IFieldMergingCallback as IFieldMergingCallback;
use com\aspose\words\FieldMergingArgs as FieldMergingArgs;
use com\aspose\words\DocumentBuilder as DocumentBuilder;
use com\aspose\words\TextFormFieldType as TextFormFieldType;
use com\aspose\words\ImageFieldMergingArgs as ImageFieldMergingArgs;
use java\lang\Boolean as Boolean;
use java\text\MessageFormat as MessageFormat;


class MailMergeFormFields {

    public static function main()
    {

        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/mailmergeandreporting/mailmergeformfields/data/";
        $doc = new Document($dataDir . "Template.doc");
        $doc->getMailMerge()->setFieldMergingCallback(new HandleMergeField());

        $fieldNames = array("RecipientName","SenderName","FaxNumber","PhoneNumber","Subject","Body","Urgent","ForReview","PleaseComment");
        $fieldValues = array("Josh","Jenny","123456789","","Hello","Test Pakistan 1", true, false, true);

        $doc->getMailMerge()->execute($fieldNames,$fieldValues);

        $doc->save($dataDir . "Template Out.doc");

    }

}

class HandleMergeField extends IFieldMergingCallback {

    private $mBuilder = null;
    private $mF = null;
    private $tffT = null;

    function __construct()
    {
        $this->mBuilder = new DocumentBuilder();
        $this->mF = new MessageFormat();
        $this->tffT = new TextFormFieldType();
    }

    public function fieldMerging(FieldMergingArgs $e)
    {

        if($this->mBuilder == null)
            $this->mBuilder = new DocumentBuilder($e->getDocument());

        if($e->getFieldValue() instanceof Boolean )
        {
            $this->mBuilder->moveToMergeField($e->getFieldName());
            $checkBoxName = $this->mF->format("{0}{1}",$e->getFieldName(),$e->getRecordIndex());

            $this->mBuilder->insertCheckBox($checkBoxName, (Boolean) $e->getFieldValue(), 0);

            return;
        }

        if(java_values($e->getFieldName()) == 'Subject')
        {
            $this->mBuilder->moveToMergeField($e->getFieldName());
            $textInputName = $this->mF->format("{0}{1}", $e->getFieldName(), $e->getRecordIndex());
            $this->mBuilder->insertTextInput($textInputName, $this->tffT->REGULAR, "", $e->getFieldValue(), 0 );
        }

    }

    public function imageFieldMerging(ImageFieldMergingArgs $args)
    {

    }

}

