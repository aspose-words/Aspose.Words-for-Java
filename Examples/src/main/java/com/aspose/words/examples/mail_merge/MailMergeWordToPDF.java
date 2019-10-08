package com.aspose.words.examples.mail_merge;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.linq.BubbleChart;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.regex.Pattern;

public class MailMergeWordToPDF {

    public static void main(String[] args) throws Exception {
        // TODO Auto-generated method stub
        //ExStart: MailMergeWordToPDF
        String dataDir = Utils.getDataDir(BubbleChart.class);

        String fileName = "Converted.pdf";
        com.aspose.pdf.Document pdfDoc = new com.aspose.pdf.Document(dataDir + fileName);

        ByteArrayOutputStream baOs = new ByteArrayOutputStream();
        //Converting PDF document to word document
        pdfDoc.save(baOs, com.aspose.pdf.SaveFormat.DocX);

        String[] referenceFields = {"FIRSTNAME", "MEMFIRST"};
        Object[] referenceValues = {"AAA", "BBB"};

        Document document = mailMergeTemplate(baOs.toByteArray(), referenceFields, referenceValues);
        document.save(dataDir + "Saved.pdf");
        //ExEnd: MailMergeWordToPDF
    }

    //ExStart: mailMergeTemplate
    private static Document mailMergeTemplate(byte[] templateFile, String[] referenceFields, Object[] referenceValues) throws Exception {
        Document doc = null;
        try {
            doc = new Document(new ByteArrayInputStream(templateFile));

            FindReplaceOptions opts = new FindReplaceOptions();
            opts.setFindWholeWordsOnly(false);
            opts.setReplacingCallback(new ReplaceEvaluatorFindAndInsertMergefield());

            doc.getRange().replace(Pattern.compile("«(.*?)»"), "", opts);

            doc.getMailMerge().setFieldMergingCallback(new HandleMergeFields());
            doc.getMailMerge().execute(referenceFields, referenceValues);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return doc;
    }
    //ExEnd: mailMergeTemplate

    //ExStart: HandleMergeFields

    /**
     * This is called when mail merge engine encounters plain text (non image) merge
     */
    static class HandleMergeFields implements IFieldMergingCallback {
        public void fieldMerging(FieldMergingArgs args) throws Exception {
            System.out.println("Mail merge for field : " + args.getFieldName() + " & Value : " + args.getFieldValue());
        }

        /**
         * This is called when mail merge engine encounters Image:XXX merge
         * field in the document. You have a chance to return an Image object,
         * file name or a stream that contains the image.
         */
        public void imageFieldMerging(ImageFieldMergingArgs e) throws Exception {
            System.out.println("Mail merge for field : " + e.getFieldName() + " & Value : " + e.getFieldValue());
        }
    }
    //ExEnd: HandleMergeFields

    //ExStart: ReplaceEvaluatorFindAndInsertMergefield
    static class ReplaceEvaluatorFindAndInsertMergefield implements IReplacingCallback {

        public int replacing(ReplacingArgs e) throws Exception {
            // This is a Run node that contains either the beginning or the complete match.
            Node currentNode = e.getMatchNode();

            // The first (and may be the only) run can contain text before the match,
            // in this case it is necessary to split the run.
            if (e.getMatchOffset() > 0)
                currentNode = splitRun((Run) currentNode, e.getMatchOffset());

            ArrayList runs = new ArrayList();

            // Find all runs that contain parts of the match string.
            int remainingLength = e.getMatch().group().length();
            while ((remainingLength > 0) && (currentNode != null) && (currentNode.getText().length() <= remainingLength)) {
                runs.add(currentNode);
                remainingLength = remainingLength - currentNode.getText().length();

                // Select the next Run node.
                // Have to loop because there could be other nodes such as BookmarkStart etc.
                do {
                    currentNode = currentNode.getNextSibling();
                } while ((currentNode != null) && (currentNode.getNodeType() != NodeType.RUN));
            }

            // Split the last run that contains the match if there is any text left.
            if ((currentNode != null) && (remainingLength > 0)) {
                splitRun((Run) currentNode, remainingLength);
                runs.add(currentNode);
            }

            //Change static text to real merge fields.
            DocumentBuilder builder = new DocumentBuilder((Document) e.getMatchNode().getDocument());
            builder.moveTo((Run) runs.get(runs.size() - 1));
            builder.insertField("MERGEFIELD \"" + e.getMatch().group(1) + "\"");

            for (Run run : (Iterable<Run>) runs)
                run.remove();

            // Signal to the replace engine to do nothing because we have already done all what we wanted.
            return ReplaceAction.SKIP;
        }

        /**
         * Splits text of the specified run into two runs. Inserts the new run just
         * after the specified run.
         */
        private Run splitRun(Run run, int position) throws Exception {
            Run afterRun = (Run) run.deepClone(true);
            afterRun.setText(run.getText().substring(position));
            run.setText(run.getText().substring((0), (0) + (position)));
            run.getParentNode().insertAfter(afterRun, run);
            return afterRun;
        }
    }
    //ExEnd: ReplaceEvaluatorFindAndInsertMergefield
}
