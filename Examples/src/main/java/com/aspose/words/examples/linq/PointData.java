package com.aspose.words.examples.linq;

public class PointData {
    public String Time;
    public int Flow;
    public int Rainfall;
    
    public PointData (String Time, int Flow, int RainFall){
    	this.Time = Time;
    	this.Flow = Flow;
    	this.Rainfall = RainFall;
    }
    
    public final String getTime() {
        return Time;
    }

    public final int getFlow() {
        return Flow;
    }
    
    public final int getTRainfall() {
        return Rainfall;
    }
}
