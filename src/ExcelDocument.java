import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDocument {
	private XSSFWorkbook workbook;
	private XSSFSheet activeSheet;
	
	public ExcelDocument() throws IOException {
		this.workbook = new XSSFWorkbook();
	}
	
	public ExcelDocument(String filePath) throws IOException {
		this.workbook = new XSSFWorkbook(filePath);
	}
	
	public void write(String filePath) {
		FileOutputStream fileOutputStream = null;
		try {
			fileOutputStream = new FileOutputStream(filePath);
			workbook.write(fileOutputStream);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if(fileOutputStream != null) {
				try {
					fileOutputStream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
	
	public void setColumnWidth(int columnIndex, int width) {
		activeSheet.setColumnWidth(columnIndex, width);
	}
	
	public void setValue(int rowIndex, int cellIndex, String value) {
		Cell cell = createCell(rowIndex, cellIndex);
		cell.setCellValue(value);
	}
	
	public void setValue(int rowIndex, int cellIndex, String value, ExcelStyle style) {
		Cell cell = createCell(rowIndex, cellIndex);
		cell.setCellStyle(convertStyle(style));
		cell.setCellValue(value);
	}
	
	public void setValues(int rowIndex, String[] values) {
		for (int i = 0; i < values.length; i++) {
			setValue(rowIndex, i, values[i]);
		}
	}
	
	public void addValues(String[] values) {
		activeSheet.createRow(activeSheet.getLastRowNum()+1);
		for (int i = 0; i < values.length; i++) {
			setValue(activeSheet.getLastRowNum(), i, values[i]);
		}
	}
	
	public void setValues(int rowIndex, String[] values, ExcelStyle style) {
		for (int i = 0; i < values.length; i++) {
			setValue(rowIndex, i, values[i], style);
		}
	}
	
	public void addValues(String[] values, ExcelStyle style) {
		activeSheet.createRow(activeSheet.getLastRowNum()+1);
		for (int i = 0; i < values.length; i++) {
			setValue(activeSheet.getLastRowNum(), i, values[i], style);
		}
	}
	
	/*public void addData(String[] data, ExcelStyle style) {
		int newIndex = 1;
		if(activeSheet.getLastRowNum() > 0) {
			newIndex = activeSheet.getLastRowNum() + 1;
		}
		Row newRow = activeSheet.createRow(newIndex);
		for (int i = 0; i < data.length; i++) {
			Cell newCell = addStyledCell(newRow, i, convertStyle(style));
			newCell.setCellValue(data[i]==null?"":data[i]);
		}
	}*/
	
	public void selectSheetByIndex(int sheetIndex) {
		activeSheet = workbook.getSheetAt(sheetIndex);
	}
	
	public void selectSheetByTitle(String sheetTitle) {
		activeSheet = workbook.getSheet(sheetTitle);
	}
	
	public void createSheet(String sheetTitle) {
		workbook.createSheet(sheetTitle);
	}
	
	private Cell getCell(int rowIndex, int cellIndex) {
		Row row = activeSheet.getRow(rowIndex);
		if(row == null) {
			return null;
		}
		return row.getCell(cellIndex);
	}
	
	public String getStringValue(int rowIndex, int cellIndex)
	{
		Cell cell = getCell(rowIndex, cellIndex);
		if(cell == null) {
			return null;
		}
		return cell.getStringCellValue();
	}
	
	public String getStringValue(int rowIndex, int cellIndex, String defaultValue)
	{
		String result = getStringValue(rowIndex, cellIndex);
		if(result == null) {
			return defaultValue;
		}
		return result;
	}
	
	private Cell createCell(int rowIndex, int cellIndex) {
		Row row = activeSheet.getRow(rowIndex);
		
		if(row == null) {
			row = activeSheet.createRow(rowIndex);
		}
		
		Cell cell = row.getCell(cellIndex);
		
		if(cell == null) {
			cell = row.createCell(cellIndex);
		}
		
		return cell;
	}
	
	public int getRowsCount() {
		if(activeSheet == null) {
			return -1;
		}
		return activeSheet.getLastRowNum();
	}
	
	public CellStyle getCellStyle(int rowIndex, int cellIndex) {
		Cell cell = getCell(rowIndex, cellIndex);
		if(cell == null) {
			return null;
		}
		return cell.getCellStyle();
	}
	
	private CellStyle convertStyle(ExcelStyle st) {
		XSSFCellStyle c = workbook.createCellStyle();
		
		if(st.getFont() != null) {
			Font f = workbook.createFont();
			if(st.getFont().isBold()) {
				f.setBold(true);
			}
			if(st.getFont().isItalic()) {
				f.setItalic(true);
			}
			c.setFont(f);
		}
		
		if(st.getBackgroundColor() != null) {
			c.setFillForegroundColor(new XSSFColor(st.getBackgroundColor()));
			c.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		}
		return c;
	}
}
