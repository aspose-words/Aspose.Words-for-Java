package ApiExamples.TestData.TestBuilders;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.System.DateTime;
import ApiExamples.TestData.TestClasses.NumericTestClass;


public class NumericTestBuilder
{
    private int? mValue1;
    private double mValue2;
    private int mValue3;
    private int? mValue4;
    private boolean mLogical;
    private DateTime mDate;

    public NumericTestBuilder()
    {
        mValue1 = 1;
        mValue2 = 1.0;
        mValue3 = 1;
        mValue4 = 1;
        mLogical = false;
        mDate = new DateTime(2018, 1, 1);
    }

    public NumericTestBuilder withValuesAndDate(int? value1, double value2, int value3, int? value4,
        DateTime dateTime)
    {
        mValue1 = value1;
        mValue2 = value2;
        mValue3 = value3;
        mValue4 = value4;
        mDate = dateTime;
        return this;
    }

    public NumericTestBuilder withValuesAndLogical(int? value1, double value2, int value3, int? value4,
        boolean logical)
    {
        mValue1 = value1;
        mValue2 = value2;
        mValue3 = value3;
        mValue4 = value4;
        mLogical = logical;
        return this;
    }

    public NumericTestClass build()
    {
        return new NumericTestClass(mValue1, mValue2, mValue3, mValue4, mLogical, mDate);
    }
}
