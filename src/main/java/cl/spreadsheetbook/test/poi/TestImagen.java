package cl.spreadsheetbook.test.poi;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Shape;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestImagen {

	public static void main(String[] args) {  

		try {

			Workbook wb = new XSSFWorkbook();
			Sheet sheet = wb.createSheet("./files/My Sample Excel");

			//FileInputStream obtains input bytes from the image file
			InputStream inputStream = new FileInputStream("logo.jpeg");
			//Get the contents of an InputStream as a byte[].
			byte[] bytes = IOUtils.toByteArray(inputStream);
			//Adds a picture to the workbook
			//int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
			int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
			//close the input stream
			inputStream.close();

			//Returns an object that handles instantiating concrete classes
			CreationHelper helper = wb.getCreationHelper();

			//Creates the top-level drawing patriarch.
			Drawing<Shape> drawing = (Drawing<Shape>) sheet.createDrawingPatriarch();

			//Create an anchor that is attached to the worksheet
			ClientAnchor anchor = helper.createClientAnchor();
			//set top-left corner for the image
			anchor.setCol1(1);
			anchor.setRow1(10);
			

			//Creates a picture
			Picture pict = drawing.createPicture(anchor, pictureIdx);
			//Reset the image to the original size
			pict.resize();

			//Write the Excel file
			FileOutputStream fileOut = null;
			fileOut = new FileOutputStream("test-imagen.xls");
			wb.write(fileOut);
			fileOut.close();

		}
		catch (Exception e) {
			System.out.println(e);
		}
	}  
}
