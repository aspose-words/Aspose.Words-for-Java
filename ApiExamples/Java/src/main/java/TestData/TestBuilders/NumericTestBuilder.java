package TestData.TestBuilders;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import TestData.TestClasses.NumericTestClass;

import java.time.LocalDate;

public class NumericTestBuilder {
    private Integer mValue1;
    private double mValue2;
    private Integer mValue3;
    private Integer mValue4;
    private boolean mLogical;
    private LocalDate mDate;

    public NumericTestBuilder() {
        mValue1 = 1;
        mValue2 = 1.0;
        mValue3 = 1;
        mValue4 = 1;
        mLogical = false;
        mDate = LocalDate.of(2018, 1, 1);
    }

    public NumericTestBuilder withValuesAndDate(final Integer value1, final double value2, final Integer value3,
                                                final Integer value4, final LocalDate dateTime) {
        mValue1 = value1;
        mValue2 = value2;
        mValue3 = value3;
        mValue4 = value4;
        mDate = dateTime;
        return this;
    }

    public NumericTestBuilder withValuesAndLogical(final Integer value1, final double value2, final Integer value3,
                                                   final Integer value4, final boolean logical) {
        mValue1 = value1;
        mValue2 = value2;
        mValue3 = value3;
        mValue4 = value4;
        mLogical = logical;
        return this;
    }

    public NumericTestClass build() {
        return new NumericTestClass(mValue1, mValue2, mValue3, mValue4, mLogical, mDate);
    }
}
