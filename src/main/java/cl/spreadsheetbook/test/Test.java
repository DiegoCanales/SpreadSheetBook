package cl.spreadsheetbook.test;

import cl.spreadsheetbook.utils.Border;
import cl.spreadsheetbook.utils.Font;
import cl.spreadsheetbook.utils.SpreadSheetBook;

public class Test {

	public static void main(String[] args) {

		// crear libro
		SpreadSheetBook book = new SpreadSheetBook("./files", "test");
		
		// añadir valor a celda
		book.addInCell(0, 0, "test !!");
		
		// estilo celda
		book.addFontStyle(0, 0, 
				new Font(true, (short)20, "red"));
		
		// imagen
		book.insertImage(10, 0, "logo.jpeg");
		
		// borde
		book.setBorderStyle(0, 0, 
				new Border(true, true, true, true, "thin", "black")
				);
	}

}
