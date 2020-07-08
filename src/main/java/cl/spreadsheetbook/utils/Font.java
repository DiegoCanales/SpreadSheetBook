package cl.spreadsheetbook.utils;

import java.util.HashMap;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFFont;

public class Font {
	
	public boolean bold;
	public short fontSize;
	public short color;
	public int alignment;
	
	// Aliniamiento
	public static int ALIGNMENT_CENTER = 0;
	public static int ALIGNMENT_LEFT = -1;
	public static int ALIGNMENT_RIGHT = 1;
	
	private HashMap<Integer, HorizontalAlignment> alignments = new HashMap<Integer, HorizontalAlignment>(){{
		put(Font.ALIGNMENT_CENTER,HorizontalAlignment.CENTER);
		put(Font.ALIGNMENT_LEFT,HorizontalAlignment.LEFT);
		put(Font.ALIGNMENT_RIGHT,HorizontalAlignment.RIGHT);
		
	}};
	
	public Font(boolean bold, short fontSize, short color, int alignment) {
		super();
		this.bold = bold;
		this.fontSize = fontSize;
		this.color = color;
		this.alignment = alignment;
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
		cellFont.setColor(this.color);
	}
	
	public void toCellStyle(CellStyle style) {
		style.setAlignment(this.getAlignment(alignment));
	}

}
