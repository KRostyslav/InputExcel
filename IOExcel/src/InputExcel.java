import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class InputExcel {

	public static void main(String args[]) {
		new InputExcel().WExcel();
	}

	public void WExcel() {

		Map<String, Object[]> data = new HashMap<String, Object[]>();
		data.put("1", new Object[] { "Emp No.", "Name", "Surname", "Age" });
		data.put("2", new Object[] { 1d, "Aaaa", "Kkkkkk", 24d });
		data.put("3", new Object[] { 2d, "Bbbb", "Mmmmm", 36d });
		data.put("4", new Object[] { 3d, "Ccc", "Nnnnnnn", 30d });
		data.put("5", new Object[] { 4d, "Ddddd", "Llll", 41d });

		Set<String> keyset = data.keySet();
		int rownum = 0;

		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("TableSheet");

		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArray = data.get(key);
			int cellnum = 0;
			for (Object obj : objArray) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof Date)
					cell.setCellValue((Date) obj);
				else if (obj instanceof Boolean)
					cell.setCellValue((Boolean) obj);
				else if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Double)
					cell.setCellValue((Double) obj);
			}
		}

		try {
			FileOutputStream out = new FileOutputStream(new File(
					"Excel_from_java.xls"));
			workbook.write(out);
			out.close();
			System.out.println("Excel written successfully...");

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
