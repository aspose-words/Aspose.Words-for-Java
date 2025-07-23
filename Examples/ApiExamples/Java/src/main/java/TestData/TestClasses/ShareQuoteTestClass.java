package TestData.TestClasses;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

public class ShareQuoteTestClass {
    private int mDate;
    private int mVolume;
    private double mOpen;
    private double mHigh;
    private double mLow;
    private double mClose;

    public int getDate()
    {
        return mDate;
    }
    public int getVolume()
    {
        return mVolume;
    }

    public double getOpen()
    {
        return mOpen;
    }

    public double getHigh()
    {
        return mHigh;
    }

    public double getLow()
    {
        return mLow;
    }

    public double getClose()
    {
        return mClose;
    }

    public ShareQuoteTestClass(int date, int volume, double open, double high, double low, double close) {
        mDate = date;
        mVolume = volume;
        mOpen = open;
        mHigh = high;
        mLow = low;
        mClose = close;
    }

    public String getColor() {
        return (mOpen < mClose) ? "#1B9629" : "#96002C";
    }
}

