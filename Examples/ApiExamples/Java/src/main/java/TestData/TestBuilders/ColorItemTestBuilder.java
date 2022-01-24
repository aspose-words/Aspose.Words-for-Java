package TestData.TestBuilders;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import TestData.TestClasses.ColorItemTestClass;

import java.awt.*;

public class ColorItemTestBuilder {
    private String mName;
    private Color mColor;
    private int mColorCode;
    private double mValue1;
    private double mValue2;
    private double mValue3;

    public ColorItemTestBuilder() {
        mName = "DefaultName";
        mColor = Color.BLACK;
        mColorCode = Color.BLACK.getRGB();
        mValue1 = 1.0;
        mValue2 = 1.0;
        mValue3 = 1.0;
    }

    public ColorItemTestBuilder withColor(final String name, final Color color) {
        mName = name;
        mColor = color;
        return this;
    }

    public ColorItemTestBuilder withColorCode(final String name, final int colorCode) {
        mName = name;
        mColorCode = colorCode;
        return this;
    }

    public ColorItemTestBuilder withColorAndValues(final String name, final Color color, final double value1,
                                                   final double value2, final double value3) {
        mName = name;
        mColor = color;
        mValue1 = value1;
        mValue2 = value2;
        mValue3 = value3;
        return this;
    }

    public ColorItemTestBuilder withColorCodeAndValues(final String name, final int colorCode, final double value1,
                                                       final double value2, final double value3) {
        mName = name;
        mColorCode = colorCode;
        mValue1 = value1;
        mValue2 = value2;
        mValue3 = value3;
        return this;
    }

    public ColorItemTestClass build() {
        return new ColorItemTestClass(mName, mColor, mColorCode, mValue1, mValue2, mValue3);
    }
}
