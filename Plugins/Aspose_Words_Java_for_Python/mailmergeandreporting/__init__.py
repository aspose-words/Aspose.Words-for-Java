__author__ = 'fahadadeel'
import jpype

class MailMergeFormFields:

    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Document = jpype.JClass("com.aspose.words.Document")

    def main(self):

        # Load the template document.
        doc = self.Document(self.dataDir + "Template.doc")

        # Setup mail merge event handler to do the custom work.

        c = HandleMergeField()

        proxy = jpype.JProxy("com.aspose.words.IFieldMergingCallback", inst=c)

        doc.getMailMerge().setFieldMergingCallback(proxy)

        # This is the data for mail merge.
        fieldNames = ["RecipientName", "SenderName", "FaxNumber", "PhoneNumber",
            "Subject", "Body", "Urgent", "ForReview", "PleaseComment"]
        fieldValues = ["Josh", "Jenny", "123456789", "", "Hello",
            "Test message 1", True, False, True]

        # Execute the mail merge.
        doc.getMailMerge().execute(fieldNames, fieldValues)

        # Save the finished document.
        doc.save(self.dataDir + "Template Out.doc")

class HandleMergeField:

    def __init__(self):

        self.DocumentBuilder = jpype.JClass("com.aspose.words.DocumentBuilder")
        self.TextFormFieldType = jpype.JClass("com.aspose.words.TextFormFieldType")
        self.mBuilder = self.DocumentBuilder()

    def fieldMerging(self,e):

        if (self.mBuilder is None):
                self.mBuilder = self.DocumentBuilder(e.getDocument())

        # We decided that we want all boolean values to be output as check box form fields.
        if (e.getFieldValue() in ('True', 'False')) :
            # Move the "cursor" to the current merge field.
            self.mBuilder.moveToMergeField(e.getFieldName())

            # It is nice to give names to check boxes. Lets generate a name such as MyField21 or so.
            checkBoxName = e.getFieldName() + e.getRecordIndex()

            # Insert a check box.
            self.mBuilder.insertCheckBox(checkBoxName, e.getFieldValue(), 0)

            # Nothing else to do for this field.
            return

        # Another example, we want the Subject field to come out as text input form field.
        if ("Subject" == e.getFieldName()):

            self.mBuilder.moveToMergeField(e.getFieldName())
            textInputName = e.getFieldName() + e.getRecordIndex()
            self.mBuilder.insertTextInput(textInputName, self.TextFormFieldType.REGULAR, "", e.getFieldValue(), 0)

    def imageFieldMerging(self):

        return

