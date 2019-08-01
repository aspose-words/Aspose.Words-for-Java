package ApiExamples.TestData.TestBuilders;

// ********* THIS FILE IS AUTO PORTED *********

import java.awt.Color;
import com.aspose.ms.System.Drawing.msColor;
import ApiExamples.TestData.TestClasses.ColorItemTestClass;


public class ColorItemTestBuilder
{
    public String Name;
    public Color _Color = msColor.Empty;
    public int ColorCode;
    public double Value1;
    public double Value2;
    public double Value3;

    public ColorItemTestBuilder()
    {
        Name = "DefaultName";
        _Color = Color.BLACK;
        ColorCode = Color.BLACK.getRGB();
        Value1 = 1.0;
        Value2 = 1.0;
        Value3 = 1.0;
    }

    public ColorItemTestBuilder withColor(String name, Color color)
    {
        Name = name;
        _Color = color;
        return this;
    }

    public ColorItemTestBuilder withColorCode(String name, int colorCode)
    {
        Name = name;
        ColorCode = colorCode;
        return this;
    }

    public ColorItemTestBuilder withColorAndValues(String name, Color color, double value1, double value2,
        double value3)
    {
        Name = name;
        _Color = color;
        Value1 = value1;
        Value2 = value2;
        Value3 = value3;
        return this;
    }

    public ColorItemTestBuilder withColorCodeAndValues(String name, int colorCode, double value1, double value2,
        double value3)
    {
        Name = name;
        ColorCode = colorCode;
        Value1 = value1;
        Value2 = value2;
        Value3 = value3;
        return this;
    }

    public ColorItemTestClass build()
    {
        return new ColorItemTestClass(Name, _Color, ColorCode, Value1, Value2, Value3);
    }
}
