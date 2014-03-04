/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package loadingandsaving.loadingandsavinghtml.savehtmlandemail.java;

import javax.swing.text.MaskFormatter;
import java.text.ParseException;

/**
 * Allows input verified by a mask to be of variable length.
 */
public class VariableLengthMaskFormatter extends MaskFormatter {

    public VariableLengthMaskFormatter() {
        super();
    }

    public VariableLengthMaskFormatter(String mask) throws ParseException {
        super(mask);
    }

    public Object stringToValue(String value) throws ParseException {
        Object rv;
        String mask = getMask();

        if (mask != null) {
            setMask( getMaskForString(mask, value));
            rv = super.stringToValue(value.substring( 0, getMask().length()));
            setMask(mask);
        }
        else {
            rv = super.stringToValue(value);
        }

        return rv;
    }


    protected String getMaskForString( String mask, String value ) {

        StringBuffer sb = new StringBuffer();
        int maskLength = mask.length();
        char placeHolder = getPlaceholderCharacter();

        for (int k = 0, size = value.length(); k < size && k < maskLength ; k++) {

            if ( placeHolder == value.charAt( k ) )
                break;

            sb.append( mask.charAt( k ) );
        }
        return sb.toString();

    }
}