package cl.spreadsheetbook.utils;

import java.util.HashMap;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

public class Border {
	public boolean top;
	public boolean bottom;
	public boolean left;
	public boolean right;
	public String style;
	public String color;
	
	private HashMap<String, BorderStyle> styles = new HashMap<String,BorderStyle>(){{
		put("thin",BorderStyle.THIN);
		put("thick",BorderStyle.THICK);
		put("none",BorderStyle.NONE);
	}};
	
	private  HashMap<String, Short> colors = new HashMap<String,Short>(){{
		put("black",IndexedColors.BLACK.getIndex());
		put("red",IndexedColors.RED.getIndex());
		put("green",IndexedColors.GREEN.getIndex());
		put("blue",IndexedColors.BLUE.getIndex());
	}};
	
	
	public Border() {
		
	}

	public Border(boolean top, boolean bottom, boolean left, boolean right,String style, String color) {
		super();
		this.top = top;
		this.bottom = bottom;
		this.left = left;
		this.right = right;
		this.style = style;
		this.color = color;
	}
	
	public void addBorderStyle(CellStyle style) {
		// border
		BorderStyle bs = styles.containsKey(this.style) ? styles.get(this.style) : styles.get("none");
		if(top)
			style.setBorderTop(bs);
		if(bottom)
			style.setBorderBottom(bs);
		if(left)
			style.setBorderLeft(bs);
		if(right)
			style.setBorderRight(bs);
		
		//border color
		short color = colors.containsKey(this.color) ? colors.get(this.color) : colors.get("black");
		style.setTopBorderColor(color);
		style.setBottomBorderColor(color);
		style.setLeftBorderColor(color);
		style.setRightBorderColor(color);
	}
	
}
