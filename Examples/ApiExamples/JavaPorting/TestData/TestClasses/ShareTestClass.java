package ApiExamples.TestData.TestClasses;

// ********* THIS FILE IS AUTO PORTED *********

import java.text.MessageFormat;
import com.aspose.FormatterPal;
import com.aspose.ms.System.msMath;
import com.aspose.ms.System.msByte;

public class ShareTestClass
{
    ShareTestClass(String sector, String industry, String ticker, double weight, double delta)
    {
        Sector = sector;
        Industry = industry;
        Ticker = ticker;
        Weight = weight;
        Delta = delta;
    }

    public String title()
    {
        double percentValue = Delta * 100.0;
        return MessageFormat.format("{0}\r\n{1}%", Ticker, FormatterPal.doubleToStr(percentValue));
    }

    public String color()
    {
        final double FULL_COLOR_DELTA = 0.016d;
        final byte UNUSED_COLOR_CHANNEL_VALUE = (byte) 80;

        byte r = UNUSED_COLOR_CHANNEL_VALUE;
        byte g = UNUSED_COLOR_CHANNEL_VALUE;
        byte b = UNUSED_COLOR_CHANNEL_VALUE;

        int value =
            (UNUSED_COLOR_CHANNEL_VALUE & 0xFF) +
            (int)msMath.round(Math.abs(Delta) / FULL_COLOR_DELTA *
                ((msByte.MAX_VALUE & 0xFF) - (UNUSED_COLOR_CHANNEL_VALUE & 0xFF)));

        if (value > (msByte.MAX_VALUE & 0xFF))
            value = (msByte.MAX_VALUE & 0xFF);

        if (Delta < 0)
            r = (byte)value;
        else
            g = (byte)value;

        return MessageFormat.format("#{0:X2}{1:X2}{2:X2}", (r & 0xFF), (g & 0xFF), (b & 0xFF));
    }

    public String industryColor()
    {
        if ("Consumer Electronics".equals(Industry))
            return "#1B9629";
        else if ("Software - Infrastructure".equals(Industry))
            return "#6029E3";
        else if ("Semiconductors".equals(Industry))
            return "#E38529";
        else if ("Internet Content & Information".equals(Industry))
            return "#964D05";
        else if ("Entertainment".equals(Industry))
            return "#12E32B";
        else if ("Internet Retail".equals(Industry))
            return "#96002C";
        else if ("Auto Manufactures".equals(Industry))
            return "#1EE3A4";
        else if ("Credit Services".equals(Industry))
            return "#D40B70";
        else
            return "#888888";
    }

    public String Sector;
    public String Industry;
    public String Ticker;
    public double Weight;
    public double Delta;
}

