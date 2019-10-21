package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.IFieldResultFormatter;

import java.util.ArrayList;
import java.util.Date;

/**
 * Created by Home on 5/29/2017.
 */
//ExStart:FieldResultFormatter
public class FieldResultFormatter implements IFieldResultFormatter {

    private final String mNumberFormat;
    private final String mDateFormat;

    private final ArrayList mNumberFormatInvocations = new ArrayList();
    private final ArrayList mDateFormatInvocations = new ArrayList();

    public FieldResultFormatter(String numberFormat, String dateFormat) {
        mNumberFormat = numberFormat;
        mDateFormat = dateFormat;
    }

    public FieldResultFormatter() {
        mNumberFormat = null;
        mDateFormat = null;
    }

    public String format(String arg0, int arg1) {
        // TODO Auto-generated method stub
        return null;
    }

    public String format(double arg0, int arg1) {
        // TODO Auto-generated method stub
        return null;
    }

    public String formatNumeric(double value, String format) {
        // TODO Auto-generated method stub

        mNumberFormatInvocations.add(new Object[]{value, format});
        return (mNumberFormat.isEmpty() || mNumberFormat == null) ? null
                : String.format(mNumberFormat, value);
    }

    public String formatDateTime(Date value, String format, int calendarType) {
        mDateFormatInvocations
                .add(new Object[]{value, format, calendarType});

        return (mDateFormat.isEmpty() || mDateFormat == null) ? null : String
                .format(mDateFormat, value);
    }
//ExEnd:FieldResultFormatter
}
