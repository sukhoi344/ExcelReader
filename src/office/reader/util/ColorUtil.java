package office.reader.util;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

public class ColorUtil {
	
	/**
	 * Get RGB color value from colorIndex
	 * @param colorIndex
	 * @return array of three short elements (RGB).
	 */
	public static short[] getColorValues(HSSFWorkbook workBook, short colorIndex) {
		HSSFPalette palette = new HSSFWorkbook().getCustomPalette();
		HSSFColor color = palette.getColor(colorIndex);
		
		return color.getTriplet();
	}
	
	/**
	 * Get RGB color string in the format #FFFFFF from
	 * color Index
	 * @param colorIndex
	 * @return
	 */
	public static String getColorString(HSSFWorkbook workBook, short colorIndex) {
		short[] triplet = getColorValues(workBook, colorIndex);
		
		short red = triplet[0];
		short blue = triplet[1];
		short green = triplet[2];
		
		String redHex = Integer.toHexString(red & 0xffff);
		String blueHex = Integer.toHexString(blue & 0xffff);
		String greenHex = Integer.toHexString(green & 0xffff);
		
		String colorString = "#" + ((redHex.length()==1)? "0" : "") + redHex 
							   +   ((blueHex.length()==1)? "0" : "") + blueHex
							   +   ((greenHex.length()==1)? "0" : "") + greenHex;
		
		return colorString.toUpperCase();
	}
}
