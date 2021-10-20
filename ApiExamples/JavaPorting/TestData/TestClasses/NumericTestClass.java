package ApiExamples.TestData.TestClasses;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.System.DateTime;


public class NumericTestClass
{
    public int? getValue1() { return mValue1; }; public void setValue1(int? value) { mValue1 = value; };

    private int? mValue1;
    public double getValue2() { return mValue2; }; public void setValue2(double value) { mValue2 = value; };

    private double mValue2;
    public int getValue3() { return mValue3; }; public void setValue3(int value) { mValue3 = value; };

    private int mValue3;
    public int? getValue4() { return mValue4; }; public void setValue4(int? value) { mValue4 = value; };

    private int? mValue4;
    public boolean getLogical() { return mLogical; }; public void setLogical(boolean value) { mLogical = value; };

    private boolean mLogical;
    public DateTime getDate() { return mDate; }; public void setDate(DateTime value) { mDate = value; };

    private DateTime mDate;

    public NumericTestClass(int? value1, double value2, int value3, int? value4, boolean logical, DateTime dateTime)
    {
        setValue1(value1);
        setValue2(value2);
        setValue3(value3);
        setValue4(value4);
        setLogical(logical);
        setDate(dateTime);
    }

    public int sum(int value1, int value2)
    {
        int result = value1 + value2;
        return result;
    }
}
