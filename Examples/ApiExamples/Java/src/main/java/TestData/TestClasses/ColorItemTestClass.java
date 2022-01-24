package TestData.TestClasses;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import java.awt.*;

public class ColorItemTestClass {
    private String mName;
    private Color mColor;
    private int mColorCode;
    private double mValue1;
    private double mValue2;
    private double mValue3;

    public ColorItemTestClass() {
    }

    public ColorItemTestClass(final String name, final Color color, final int colorCode, final double value1,
                              final double value2, final double value3) {
        setName(name);
        setColor(color);
        setColorCode(colorCode);
        setValue1(value1);
        setValue2(value2);
        setValue3(value3);
    }

    public void setName(final String value) {
        mName = value;
    }

    public void setColor(final Color value) {
        mColor = value;
    }

    public void setColorCode(final int value) {
        mColorCode = value;
    }

    public void setValue1(final double value) {
        mValue1 = value;
    }

    public void setValue2(final double value) {
        mValue2 = value;
    }

    public void setValue3(final double value) {
        mValue3 = value;
    }

    public String getName() {
        return mName;
    }

    public Color getColor() {
        return mColor;
    }

    public int getColorCode() {
        return mColorCode;
    }

    public double getValue1() {
        return mValue1;
    }

    public double getValue2() {
        return mValue2;
    }

    public double getValue3() {
        return mValue3;
    }
}
