import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;
import java.util.StringTokenizer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * TODO To change the template for this generated type comment go to Window -
 * Preferences - Java - Code Style - Code Templates
 */
public class RemoveDuplicate {

 /**
  *  
  */
 public RemoveDuplicate() {
  super();
  // TODO Auto-generated constructor stub
 }

public static void main(String[] args) {
  Workbook workBook = null;
  Sheet sheet = null;
  FileInputStream fs = null;
  try {
   fs = new FileInputStream(new File("./Street_03001 (version 1).xlsx"));
   workBook = new XSSFWorkbook(fs);
   sheet = workBook.getSheetAt(0);
   int rows = sheet.getPhysicalNumberOfRows();
   System.out.println(rows);
   Set s=new HashSet();
   String str="";
   for (int i = 0; i < rows; i++) {
    str="";
    Row row = sheet.getRow(i);
    int columns = row.getPhysicalNumberOfCells();
    for (int j = 0; j < columns; ++j) {
     Cell cell0 = row.getCell(j);
     System.out.println(i+"   "+j+"  "+cell0);
     if(cell0 != null){int type=cell0.getCellType();
     System.out.println(type);
     if(type==0){
      double intValue= cell0.getNumericCellValue();
      str=str+String.valueOf(intValue)+",";
      }else if(type==1){
      String stringValue=cell0.getStringCellValue();
      str=str+stringValue+",";
     }
     }
    }
    str=str
    .replace(str.charAt(str
      .lastIndexOf(",")), ' ');
    s.add(str.trim());
   }
   StringTokenizer st=null;
   String result="";
   Iterator iter=s.iterator();
   
   //Create a new workbook for the output excel
         Workbook workBookOut = new XSSFWorkbook();
         
         //Create a new Sheet in the output excel workbook
         Sheet sheetOut = workBookOut.createSheet("Remove Duplicates");
         Row[] row = new Row[s.size()];
         int rowCount=0;
         int cellCount=0;
   while(iter.hasNext()){
    cellCount=0;
    row[rowCount] = sheetOut.createRow(rowCount);
    result=iter.next().toString();
    System.out.println(result);
    st=new StringTokenizer(result," ");
    Cell[] cell= new Cell[st.countTokens()];
    while(st.hasMoreTokens()){
     cell[cellCount]=row[rowCount].createCell((short)cellCount);
     cell[cellCount].setCellValue(st.nextToken());
     ++cellCount;
    }
    ++rowCount;
   }
   FileOutputStream fileOut = new FileOutputStream(new File("./Street_03001 (version 1)Dup.xlsx"));
            workBookOut.write(fileOut);
            fileOut.close();
            System.out.println("done");
  } catch (IOException ioe) {
   ioe.printStackTrace();
  }
 }}

