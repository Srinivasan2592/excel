package exceldataget;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	private FileInputStream filein;
	private FileOutputStream fileout;
	private XSSFWorkbook wb;
	private XSSFSheet sh;
	private Cell cell;
	private int row;
	private CellStyle cellstyle;
	private Color myclor;
	private String excelpath;
	private Map<String, Integer> columns = new HashMap<>();
	static String Excelpath = "C:\\Users\\Srinivasan\\workspace\\Excel_testing\\jars\\Config.xlsx";
	static String SheetName = "Sheet1";

	public void setExcelFile(String Excelpath, String SheetName) throws Exception {
		try {
			File filein = new File(Excelpath);
			InputStream fis = new FileInputStream(filein);
			wb = new XSSFWorkbook(fis);
			sh = wb.getSheet(SheetName);
			this.excelpath = Excelpath;
			
			sh.getRow(row).forEach(cell->{
				columns.put(cell.getStringCellValue(), cell.getColumnIndex());
			});

		} catch (Exception e) {
			System.out.println(e.getMessage());

		}

	}

	public String getCellData(int rownum, int colnum) throws Exception {
		try {
			cell = sh.getRow(rownum).getCell(colnum);
			String CellData = null;
			switch (cell.getCellType()) {
			case STRING:
				CellData = cell.getStringCellValue();
				break;

			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					CellData = String.valueOf(cell.getDateCellValue());
				} else {
					CellData = String.valueOf((long) cell.getNumericCellValue());
				}
				break;

			case BOOLEAN:
				CellData = Boolean.toString(cell.getBooleanCellValue());
				break;

			case BLANK:
				CellData = "";
				break;
			}
			return CellData;

		} catch (Exception e) {
			return "";

		}

		 
	}

	public String getCellData(String columnname, int rownum) throws Exception {
		try {
			return getCellData(rownum, columns.get(columnname));
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return columnname;

	}

	public static void main(String[] args) throws Exception {
		Excel e = new Excel();

		e.setExcelFile(Excelpath, SheetName);
		// System.out.println(e.getCellData(,1));
/*		System.out.println(e.getCellData("username", 1));
		System.out.println(e.getCellData("password", 1));
		System.out.println(e.getCellData(1, 3));
		System.out.println(e.getCellData(1, 4));
		System.out.println(e.getCellData(1, 5));
		System.out.println(e.getCellData(1, 6));
		System.out.println(e.getCellData(1, 7));
		System.out.println(e.getCellData(1, 8));
		System.out.println(e.getCellData(1, 9));
		System.out.println(e.getCellData(1, 10));
*/
System.out.println(e.getCellData("srinivasan", 4));
	}
}
