package TestData.TestClasses;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

public class ShareTestClass {
    private String mSector;
    private String mIndustry;
    private String mTicker;
    private double mWeight;
    private double mDelta;

    public ShareTestClass(String sector, String industry, String ticker, double weight, double delta) {
        mSector = sector;
        mIndustry = industry;
        mTicker = ticker;
        mWeight = weight;
        mDelta = delta;
    }

    public String getTitle() {
        double percentValue = mDelta * 100;
        return String.format("{0}\r\n{1}%", mTicker, percentValue);
    }

    public String getColor() {
        final double fullColorDelta = 0.016d;
        final byte unusedColorChannelValue = 80;

        byte r = unusedColorChannelValue;
        byte g = unusedColorChannelValue;
        byte b = unusedColorChannelValue;

        int value =
                unusedColorChannelValue +
                        (int)Math.round(Math.abs(mDelta) / fullColorDelta *
                                (Byte.MAX_VALUE - unusedColorChannelValue));

        if (value > Byte.MAX_VALUE)
            value = Byte.MAX_VALUE;

        if (mDelta < 0)
            r = (byte)value;
        else
            g = (byte)value;

        return String.format("#{0:X2}{1:X2}{2:X2}", r, g, b);
    }

    public String getIndustryColor() {
        if (mIndustry == "Consumer Electronics")
            return "#1B9629";
        else if (mIndustry == "Software - Infrastructure")
            return "#6029E3";
        else if (mIndustry == "Semiconductors")
            return "#E38529";
        else if (mIndustry == "Internet Content & Information")
            return "#964D05";
        else if (mIndustry == "Entertainment")
            return "#12E32B";
        else if (mIndustry == "Internet Retail")
            return "#96002C";
        else if (mIndustry == "Auto Manufactures")
            return "#1EE3A4";
        else if (mIndustry == "Credit Services")
            return "#D40B70";
        else
            return "#888888";
    }
}
