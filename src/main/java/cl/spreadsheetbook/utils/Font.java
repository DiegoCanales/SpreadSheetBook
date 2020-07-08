package cl.spreadsheetbook.utils;

import java.util.HashMap;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFFont;

public class Font {
	
	public boolean bold;
	public short fontSize;
	public String color;
	public int alignment;
	
	// Aliniamiento
	public static int ALIGNMENT_CENTER = 0;
	public static int ALIGNMENT_LEFT = -1;
	public static int ALIGNMENT_RIGHT = 1;
	
	
	public Font(boolean bold, short fontSize, String color, int alignment) {
		super();
		this.bold = bold;
		this.fontSize = fontSize;
		this.color = color;
		this.alignment = alignment;
	}

	private HashMap<String, Short> colors = new HashMap<String,Short>(){{
		put("black",IndexedColors.BLACK.getIndex());
		put("red",IndexedColors.RED.getIndex());
		put("green",IndexedColors.GREEN.getIndex());
		put("blue",IndexedColors.BLUE.getIndex());
	}};
	
	private HashMap<Integer, HorizontalAlignment> alignments = new HashMap<Integer, HorizontalAlignment>(){{
		put(Font.ALIGNMENT_CENTER,HorizontalAlignment.CENTER);
		put(Font.ALIGNMENT_LEFT,HorizontalAlignment.LEFT);
		put(Font.ALIGNMENT_RIGHT,HorizontalAlignment.RIGHT);
		
	}};
			
	private short getColor(String color){
		return colors.containsKey(color) ? colors.get(color) : colors.get("black");
	}
	
	private HorizontalAlignment getAlignment(int alignment) {
		return alignments.containsKey(alignment) ? alignments.get(alignment) : alignments.get(Font.ALIGNMENT_CENTER);
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
	
	public void toCellStyle(CellStyle style) {
		style.setAlignment(this.getAlignment(alignment));
	}

}
