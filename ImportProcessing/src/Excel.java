import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Excel {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Workbook workbook = new XSSFWorkbook();
		
		Sheet sheet = workbook.createSheet("urmom");
		Row row = sheet.createRow(1);
		Cell cell = row.createCell(4);
		
		cell.setCellValue("I love u");

		try {
			FileOutputStream output = new FileOutputStream("Test2.xlsx");
			workbook.write(output);
			output.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
