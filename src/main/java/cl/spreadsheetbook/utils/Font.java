package cl.spreadsheetbook.utils;

import java.util.HashMap;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFFont;

public class Font {
	
	public boolean bold;
	public short fontSize;
	public String color;
	
	public Font(boolean bold, short fontSize, String color) {
		super();
		this.bold = bold;
		this.fontSize = fontSize;
		this.color = color;
	}

	private HashMap<String, Short> colors = new HashMap<String,Short>(){{
		put("black",IndexedColors.BLACK.getIndex());
		put("red",IndexedColors.RED.getIndex());
		put("green",IndexedColors.GREEN.getIndex());
		put("blue",IndexedColors.BLUE.getIndex());
	}};
			
	private short getColor(String color){
		return colors.containsKey(color) ? colors.get(color) : colors.get("black");
	}
	
	/**
	 * toXSSFFont: font debe ser creado anterioremente a partir del workbook
	 * @param font
	 */
	public void toXSSFFont(XSSFFont cellFont) {
		cellFont.setBold(this.bold);
		cellFont.setFontHeightInPoints(this.fontSize);
		cellFont.setColor(this.getColor(color));
	}
	

}
