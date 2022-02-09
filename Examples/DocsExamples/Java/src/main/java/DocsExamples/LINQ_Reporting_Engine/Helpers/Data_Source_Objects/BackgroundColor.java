package DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects;

import com.beust.jcommander.internal.Nullable;
import java.awt.Color;

//ExStart:Color
public class BackgroundColor
{
    private String mName;
    private Color mColor;
    private @Nullable int mColorCode;
    private @Nullable double mValue1;
    private @Nullable double mValue2;
    private @Nullable double mValue3;

    public String getName() { return mName; }
    public Color getColor() { return mColor; }
    public int getColorCode() { return mColorCode; }
    public double getValue1() { return mValue1; }
    public double getValue2() { return mValue2; }
    public double getValue3() { return mValue3; }

    public void setName(String value) { mName = value; }
    public void setColor(Color value) { mColor = value; }
    public void setColorCode(@Nullable int value) { mColorCode = value; }
    public void setValue1(@Nullable double value) { mValue1 = value; }
    public void setValue2(@Nullable double value) { mValue2 = value; }
    public void setValue3(@Nullable double value) { mValue3 = value; }
}
//ExEnd:Color
