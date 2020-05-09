package cl.dcnls.utils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SpreadSheetBook {
	
	private final String EXTENCION = ".xls";
	private String path = null;
	private String nombreDocumento = null;
	private File file = null;
	private XSSFWorkbook workbook = null;
	private XSSFSheet sheet = null;

	public SpreadSheetBook(String path, String nombreDocumento){
		this.path = path;
		this.nombreDocumento = nombreDocumento;
		open();
	}

	private File getFile(String path, String nombreDocumento){
		try {
			File pathFolder = new File(path);
			File file = new File(pathFolder,nombreDocumento+EXTENCION); //Crea el descriptor del archivo

			if(!file.exists()){ //Crea el archivo
				file.createNewFile();
			}
			return file;
		} catch (IOException e) { e.printStackTrace(); }
		return null;
	}

	private XSSFWorkbook getWorkbook(File file){
		try{
			XSSFWorkbook workbook = null;
			if(file.length() == 0){
				workbook = new XSSFWorkbook(); //Crea un nuevo workbook
			} else {
				workbook = (XSSFWorkbook) WorkbookFactory.create(file);
			}
			return workbook;
		} catch (InvalidFormatException e) { e.printStackTrace();
		} catch (FileNotFoundException e) {e.printStackTrace();
		} catch (IOException e) { e.printStackTrace();
		}
		return null;
	}
	
	private XSSFSheet getSheet(XSSFWorkbook workbook){
		XSSFSheet sheet;
		if(workbook.getNumberOfSheets() == 0)
			sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName(this.nombreDocumento));
		else
			sheet = workbook.getSheetAt(0);
		return sheet;
	}
	
	private void open(){
		this.file = getFile(this.path, this.nombreDocumento);
		this.workbook = getWorkbook(this.file);
		this.sheet = getSheet(this.workbook);
	}
	
	public void save(){
		try {
			File newFile = getFile(this.path, this.nombreDocumento+"_temp");
			FileOutputStream fos = new FileOutputStream(newFile);
			workbook.write(fos);
			fos.close();
			file.delete();
			newFile.renameTo(file);
			
		} catch (FileNotFoundException e) { e.printStackTrace();
		} catch (IOException e) { e.printStackTrace();
		}
	}
	
	public void delete(){
		this.file.delete();
		this.workbook = null;
		this.sheet = null;
		this.file = null;
	}

	/**
	 * isEmpty: Determina si la hoja no tiene ningun registro.
	 * @return
	 */
	public boolean isEmpty(){
		boolean noRows = getFirstEmptyRow(0) == 0;
		boolean noColumns = getFirstEmptyColumn(0) == 0;
		if(noRows && noColumns)
			return true;
		return false;
	}

	/**
	 * setActiveSheet: Selecciona una hoja.
	 * @param sheetIndex
	 */
	public void setActiveSheet(int sheetIndex){
		this.sheet = workbook.getSheetAt(sheetIndex);
	}

	/**
	 * setActiveSheet: Selecciona una hoja.
	 * @param sheetName
	 */
	public void setActiveSheet(String sheetName){
		this.sheet = workbook.getSheet(sheetName);
	}
	
	
	/**
	 * addInACell: Agrega un registro en una posicion especifica.
	 * @param rowIndex
	 * @param columnIndex
	 * @param registro
	 */
	public void addInCell(int rowIndex, int columnIndex, Object registro){
		try {
			XSSFRow row = sheet.getRow(rowIndex);
			if(row == null)
				row = sheet.createRow(rowIndex);
			XSSFCell cell = row.getCell(columnIndex);
			if(cell == null)
				cell = row.createCell(columnIndex);

			String aux = registro.getClass().getName();
			if(aux.equals(String.class.getName())) //String
				cell.setCellValue((String)registro);
			else if(aux.equals(Integer.class.getName())) //Interger
				cell.setCellValue((Integer)registro);
			else if(aux.equals(Long.class.getName())){ //Long
				Long reg = new Long((Long) registro);
				cell.setCellValue(reg.doubleValue());
			}else if(aux.equals(Double.class.getName())) //Double
				cell.setCellValue((Double)registro);
			else if(aux.equals(Boolean.class.getName())) //Boolean
				cell.setCellValue((Boolean)registro);

		} catch (EncryptedDocumentException e) { e.printStackTrace(); }
		save();
	}
	
	public void deleteCell(int rowIndex, int columnIndex){
		XSSFRow row = sheet.getRow(rowIndex);
		if(row == null)
			return;
		XSSFCell cell  = row.getCell(columnIndex);
		if(cell == null || cell.getCellTypeEnum() == CellType.BLANK){
			return;
		}else{
			cell.setCellValue("");
		}
		save();
	}
	
	/**
	 * getFirstEmptyRow: Retorna el indice de primera columna vacia.
	 * @param columnIndex
	 * @return firstEmptyRow
	 */
	public int getFirstEmptyRow(int columnIndex){
		int rowIndex = 0;
		while(true){
			XSSFRow row = sheet.getRow(rowIndex);
			if(row == null)
				row = sheet.createRow(rowIndex);
			XSSFCell cell  = row.getCell(columnIndex);
			if(cell == null || cell.getCellTypeEnum() == CellType.BLANK || getCellValue(rowIndex, columnIndex).equals("")){
				cell  = row.createCell(columnIndex);
				return rowIndex;
			}else
				rowIndex++;
		}
	}
	
	/**
	 * getFirstEmptyColumn: Retorna el indice de primera columna vacia.
	 * @param rowIndex
	 * @return firstEmptyColumn
	 */
	public int getFirstEmptyColumn(int rowIndex){
		int columnIndex = 0;
		while(true){
			XSSFRow row = sheet.getRow(rowIndex);
			if(row == null)
				row = sheet.createRow(rowIndex);
			XSSFCell cell  = row.getCell(columnIndex);
			if(cell == null || cell.getCellTypeEnum() == CellType.BLANK || getCellValue(rowIndex, columnIndex).equals("")){
				cell  = row.createCell(columnIndex);
				return columnIndex;
			}else
				columnIndex++;
		}
	}
	
	/**
	 * addInRow: Agrega un registro en la primera celda vacia de una fila.
	 * @param indexRow
	 * @param registro
	 */
	public void addInRow(int indexRow, Object registro){
		addInCell(indexRow, getFirstEmptyColumn(indexRow), registro);
	}
	
	/**
	 * addInColumn: Agrega un registro en la primera celda vacia de una columna.
	 * @param columnIndex
	 * @param registro
	 */
	public void addInColumn(int columnIndex, Object registro){
		addInCell(getFirstEmptyRow(columnIndex), columnIndex, registro);
	}
	
	/**
	 * getCellValue: Obtiene el valor de una celda espefifica.
	 * @param rowIndex
	 * @param columnIndex
	 * @return object
	 */
	public Object getCellValue(int rowIndex, int columnIndex){
		XSSFRow row = sheet.getRow(rowIndex);
		if(row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell  = row.getCell(columnIndex);
		if(cell == null)
			return null;

		if(cell.getCellTypeEnum() == CellType.STRING)
			return cell.getStringCellValue();
		if(cell.getCellTypeEnum() == CellType.NUMERIC){
			return cell.getRawValue(); //String valor crudo de la celda
			/*
			double numero = cell.getNumericCellValue();
			if(numero-(long)numero == 0){
				return (int)numero;
			}else
				return cell.getNumericCellValue();
			*/			
		}
		if(cell.getCellTypeEnum() == CellType.BOOLEAN)
			return cell.getBooleanCellValue();
		if(cell.getCellTypeEnum() == CellType.BLANK)
			return null;
		return null;
	}
	
	
	/**
	 * isCellNull: Identifica si la celda tiene un valor nulo.
	 * @param rowIndex
	 * @param columnIndex
	 * @return el valor booleano correspondiente a si la celda es nula.
	 */
	
	public boolean isCellNull(int rowIndex, int columnIndex){
		Object aux = getCellValue(rowIndex, columnIndex);
		if(aux == null || aux.equals(""))
			return true;
		else
			return false;
	}
	
	/**
	 * searchInRow: Busca un registro en una fila. Retorna el indice de la columna donde encuentra el registro y -1 si no lo encuentra.
	 * @param rowIndex
	 * @param registro
	 * @return columnIndex
	 */
	public int searchInRow(int rowIndex, Object registro){
		XSSFRow row = sheet.getRow(rowIndex);
		int columnIndex = 0;
		XSSFCell cell  = row.getCell(columnIndex);
		while(cell.getCellTypeEnum() != CellType.BLANK){
			String aux = registro.getClass().getName();

			//Celda String
			if(cell.getCellTypeEnum() == CellType.STRING && aux.equals(String.class.getName())){
				if(cell.getStringCellValue().equals((String)registro))
					return cell.getColumnIndex();
			}
			//Celda Numerica
			boolean registroIsNumeric = aux.equals(Integer.class.getName()) || aux.equals(Double.class.getName());
			if(cell.getCellTypeEnum() == CellType.NUMERIC &&  registroIsNumeric ){
				if(cell.getNumericCellValue() == (Integer)registro)
					return cell.getColumnIndex();
			}

			//Celda Booleana
			if(cell.getCellTypeEnum() == CellType.BOOLEAN && aux.equals(Boolean.class.getName())){
				if(cell.getBooleanCellValue() == (Boolean)registro)
					return cell.getColumnIndex();
			}

			//Si no es igual, busca en la siguiente columna
			columnIndex++;
			cell = row.getCell(columnIndex);
		}
		return (Integer) null;
	}

	/**
	 * searchInColumn: Busca un registro en una columna. Retorna el indice de la fila donde se encuentra el registro y -1 si no lo encuentra. 
	 * @param columnIndex
	 * @param registro
	 * @return rowIndex
	 */
	public int searchInColumn(int columnIndex, Object registro){
		int rowIndex = 0;
		XSSFRow row = sheet.getRow(rowIndex);
		XSSFCell cell  = row.getCell(columnIndex);
		while(cell.getCellTypeEnum() != CellType.BLANK){
			String aux = registro.getClass().getName();

			//Celda String
			if(cell.getCellTypeEnum() == CellType.STRING && aux.equals(String.class.getName())){
				if(cell.getStringCellValue().equals((String)registro))
					return cell.getColumnIndex();
			}
			//Celda Numerica
			boolean registroIsNumeric = aux.equals(Integer.class.getName()) || aux.equals(Double.class.getName());
			if(cell.getCellTypeEnum() == CellType.NUMERIC &&  registroIsNumeric ){
				if(cell.getNumericCellValue() == (Integer)registro)
					return cell.getColumnIndex();
			}

			//Celda Booleana
			if(cell.getCellTypeEnum() == CellType.BOOLEAN && aux.equals(Boolean.class.getName())){
				if(cell.getBooleanCellValue() == (Boolean)registro)
					return cell.getColumnIndex();
			}
			
			//Si no es igual, busca en la siguiente fila
			rowIndex++;
			row = sheet.getRow(rowIndex);
			cell = row.getCell(columnIndex);
		}
		return (Integer) null;
	}
	
}
