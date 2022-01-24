package TestData.TestClasses;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import java.time.LocalDate;

public class NumericTestClass {
    private Integer mValue1;
    private double mValue2;
    private Integer mValue3;
    private Integer mValue4;
    private boolean mLogical;
    private LocalDate mDate;

    public NumericTestClass() {
    }

    public NumericTestClass(final Integer value1, final double value2, final Integer value3, final Integer value4,
                            final boolean logical, final LocalDate dateTime) {
        setValue1(value1);
        setValue2(value2);
        setValue3(value3);
        setValue4(value4);
        setLogical(logical);
        setDate(dateTime);
    }

    public void setValue1(final Integer value) {
        mValue1 = value;
    }

    public void setValue2(final double value) {
        mValue2 = value;
    }

    public void setValue3(final Integer value) {
        mValue3 = value;
    }

    public void setValue4(final Integer value) {
        mValue4 = value;
    }

    public void setLogical(final boolean value) {
        mLogical = value;
    }

    public void setDate(final LocalDate value) {
        mDate = value;
    }

    public Integer getValue1() {
        return mValue1;
    }

    public double getValue2() {
        return mValue2;
    }

    public Integer getValue3() {
        return mValue3;
    }

    public Integer getValue4() {
        return mValue4;
    }

    public boolean getLogical() {
        return mLogical;
    }

    public LocalDate getDate() {
        return mDate;
    }

    public int sum(final int value1, final int value2) {
        int result = value1 + value2;
        return result;
    }
}
