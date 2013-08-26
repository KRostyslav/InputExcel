import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

public class InputExcel {
	
	public static void main(String args[]) {
		new InputExcel().WExcel();
	}

	public void WExcel() {
		
		HSSFWorkbook workbook = new HSSFWorkbook();		
		HSSFSheet sheet = workbook.createSheet("TableSheet");
		HSSFRow row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue((String) "Java One in cell A1");

		try {
			FileOutputStream out = new FileOutputStream(new File("Excel_from_java.xls"));
			workbook.write(out);
			out.close();

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
