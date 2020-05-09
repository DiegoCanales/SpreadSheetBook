package cl.dcnls.test;

import cl.dcnls.utils.SpreadSheetBook;

public class Test {

	public static void main(String[] args) {
		SpreadSheetBook book = new SpreadSheetBook(".", "hola");
		book.addInCell(0, 0, "hola");
		

	}

}
