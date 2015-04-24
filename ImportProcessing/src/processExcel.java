import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class processExcel {
	static FileOutputStream fos;
	static Scanner x = new Scanner(System.in);
	
	public static void main(String[] args) {
		ArrayList<String> columns = new ArrayList<String>();
		File file = new File("./toImport/test");
		String newName = file.getName()+"Dup.xlsx";
		Workbook workbook = null;
		try {
			workbook = new XSSFWorkbook(new FileInputStream(file));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		Workbook workbook2 = new XSSFWorkbook();
		columns = process(workbook,workbook2,file);
		int columnsDup = Integer.parseInt(x.next());
		try {
			fos = new FileOutputStream(newName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			findDup(workbook,workbook2,fos,columnsDup);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public static void printSheet(String[][] sheet){
		for(int i=0; i<sheet.length/sheet[0].length; i++){
			for(int j=0; j<sheet[0].length; j++){
				System.out.print(sheet[i][j]);
			}
			System.out.println();
		}
	}
	
	public static ArrayList<String> process(Workbook workbook, Workbook workbook2, File file) {
		ArrayList<String> columns = new ArrayList<String>();

		Sheet sheet = workbook.getSheetAt(0);
		Sheet sheet2;
		try{
			sheet2 = workbook2.createSheet("Duplicates");
		}catch (IllegalArgumentException e){
			sheet2 = workbook2.getSheet("Duplicates");
		}

		int m=0;
		Row row3 = sheet2.createRow(0);
		Row row33 = sheet.getRow(0);
		System.out.println("Which column to check for duplicate");

		for(int a = 0; a<row33.getLastCellNum(); a++){
			Cell cell33 = row33.getCell(a);
			if(cell33 == null){
				row33.createCell(a);
				cell33 = row33.getCell(a);
			}
			columns.add(cell33.getStringCellValue());
			System.out.print(cell33.getStringCellValue()+" "+m+" | ");
			m++;
			Cell cell4 = row33.getCell(a);
			Cell cell5 = row3.createCell(a);
			cell4.setCellType(Cell.CELL_TYPE_STRING);
			cell5.setCellValue(cell4.getStringCellValue());
		}
		
		return columns;
	}
	public static void findDup(Workbook workbook, Workbook workbook2, FileOutputStream fos, int columnDup) throws Exception{
		long timeStart = System.currentTimeMillis();
		Sheet sheet = workbook.getSheetAt(0);
		Sheet sheet2;
		try{
			sheet2 = workbook2.createSheet("Duplicates");
		}catch (IllegalArgumentException e){
			sheet2 = workbook2.getSheet("Duplicates");
		}		int j=1;
		String compare1="";
		String compare2="";
		for(int z = 1; z<sheet.getLastRowNum(); z++){
			if(z%100 == 1){
				System.out.println(z + "  time: "+(System.currentTimeMillis() - timeStart));
			}
			Row row = sheet.getRow(z);

				Cell cell = row.getCell(columnDup);
				if(cell == null){
					row.createCell(columnDup);
					cell = row.getCell(columnDup);
				}
				cell.setCellType(Cell.CELL_TYPE_STRING);

					compare2 = cell.getStringCellValue();
					//System.out.println("\n"+compare1+"  ==  "+compare2+" test");

					if(compare1.equals(compare2)){
						System.out.println(compare1+"  ==  "+compare2+" duplicate");
						Row row2 = sheet2.createRow(j);
						Row row21 = sheet2.createRow(j+1);

						j+=2;
						for(int c = 0; c<sheet.getRow(z-1).getLastCellNum(); c++){
							Cell cell2 = sheet.getRow(z-1).getCell(c);
							Cell cell3 = row2.createCell(c);
							if(cell2 == null){
								sheet.getRow(z-1).createCell(c);
								cell2 = sheet.getRow(z-1).getCell(c);
							}
							if(cell3 == null){
								row2.createCell(c);
								cell3 = row2.getCell(c);
							}
							cell2.setCellType(Cell.CELL_TYPE_STRING);
							cell3.setCellValue(cell2.getStringCellValue());
						}
						for(int b = 0; b<row.getLastCellNum(); b++){
							Cell cell21 = row.getCell(b);
							Cell cell31 = row21.createCell(b);
							if(cell21 == null){
								row.createCell(b);
								cell21 = row.getCell(b);
							}if(cell31 == null){
								row21.createCell(b);
								cell31 = row21.getCell(b);
							}
							cell21.setCellType(Cell.CELL_TYPE_STRING);
							cell31.setCellValue(cell21.getStringCellValue());
						}
						
					}
					compare1 = compare2;
		}
		System.out.println("done process");

		workbook2.write(fos);
		System.out.println("done writing");

		fos.flush();
        fos.close();
        workbook.close();
        workbook2.close();

		System.out.println("done");


	}

}
