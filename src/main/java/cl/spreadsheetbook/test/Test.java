package cl.spreadsheetbook.test;

import cl.spreadsheetbook.utils.Cell;
import cl.spreadsheetbook.utils.Color;
import cl.spreadsheetbook.utils.Font;
import cl.spreadsheetbook.utils.SpreadSheetBook;

public class Test {

	public static void main(String[] args) {

		// crear libro
		SpreadSheetBook book = new SpreadSheetBook("./files", "test");
		
		// borde
		book.addInCell(3, 0, "hola");
		book.addFontStyle(3, 0, 
				new Font(true, (short)11, Color.BLACK, Font.ALIGNMENT_CENTER));
		book.setCellStyle(3, 0, 
						new Cell(true, true, true, true, Cell.BORDER_MEDIUM, Color.GREY_25));
		
		
		book.mergeCells(0, 1, 0, 3);
		
		// añadir valor a celda
		book.addInCell(0, 0, "test !!");
		
		// estilo celda
		book.addFontStyle(0, 0, 
				new Font(true, (short)20, Color.RED, Font.ALIGNMENT_CENTER));
		
		// imagen
		//book.insertImage(10, 0, "logo.jpeg");
		
		
		
	}

}
