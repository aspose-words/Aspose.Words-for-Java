/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

package programmingwithdocuments.usingfindandreplace.replacetextwithfield.java;

import com.aspose.words.*;

import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.regex.Pattern;

public class ReplaceTextWithField
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmingwithdocuments/usingfindandreplace/replacetextwithfield/data/";

        Document doc = new Document(dataDir + "Field.ReplaceTextWithFields.doc");

        // Replace any "PlaceHolderX" instances in the document (where X is a number) with a merge field.
        Pattern regex = Pattern.compile("PlaceHolder(\\d+)", Pattern.CASE_INSENSITIVE);
        doc.getRange().replace(regex, new ReplaceTextWithFieldHandler("MERGEFIELD"), false);

        doc.save(dataDir + "Field.ReplaceTextWithFields Out.doc");
    }

    private static class ReplaceTextWithFieldHandler implements IReplacingCallback {

        public ReplaceTextWithFieldHandler(String name) {
            mFieldName = name.toUpperCase();
        }

        public int replacing(ReplacingArgs e) throws Exception {
            ArrayList runs = FindAndSplitMatchRuns(e);

            // Create DocumentBuilder which is used to insert the field.
            DocumentBuilder builder = new DocumentBuilder((Document) e.getMatchNode().getDocument());
            builder.moveTo((Run) runs.get(runs.size() - 1));

            // Insert the field into the document using the specified field type and the match text as the field name.
            // If the fields you are inserting do not require this extra parameter then it can be removed from the string below.
            builder.insertField(MessageFormat.format("{0} {1}", mFieldName, e.getMatch().group(0)));

            // Now remove all runs in the sequence.
            for (Run run : (Iterable<Run>) runs)
                run.remove();

            // Signal to the replace engine to do nothing because we have already done all what we wanted.
            return ReplaceAction.SKIP;
        }

        /**
         * Finds and splits the match runs and returns them in an ArrayList.
         */
        public ArrayList FindAndSplitMatchRuns(ReplacingArgs e) throws Exception {
            // This is a Run node that contains either the beginning or the complete match.
            Node currentNode = e.getMatchNode();

            // The first (and may be the only) run can contain text before the match,
            // in this case it is necessary to split the run.
            if (e.getMatchOffset() > 0)
                currentNode = SplitRun((Run) currentNode, e.getMatchOffset());

            // This array is used to store all nodes of the match for further removing.
            ArrayList runs = new ArrayList();

            // Find all runs that contain parts of the match string.
            int remainingLength = e.getMatch().group().length();
            while (
                    (remainingLength > 0) &&
                            (currentNode != null) &&
                            (currentNode.getText().length() <= remainingLength)) {
                runs.add(currentNode);
                remainingLength = remainingLength - currentNode.getText().length();

                // Select the next Run node.
                // Have to loop because there could be other nodes such as BookmarkStart etc.
                do {
                    currentNode = currentNode.getNextSibling();
                }
                while ((currentNode != null) && (currentNode.getNodeType() != NodeType.RUN));
            }

            // Split the last run that contains the match if there is any text left.
            if ((currentNode != null) && (remainingLength > 0)) {
                SplitRun((Run) currentNode, remainingLength);
                runs.add(currentNode);
            }

            return runs;
        }

        /**
         * Splits text of the specified run into two runs.
         * Inserts the new run just after the specified run.
         */
        private Run SplitRun(Run run, int position) throws Exception {
            Run afterRun = (Run) run.deepClone(true);
            afterRun.setText(run.getText().substring(position));
            run.setText(run.getText().substring(0, position));
            run.getParentNode().insertAfter(afterRun, run);
            return afterRun;

        }

        private String mFieldName;
    }
}




