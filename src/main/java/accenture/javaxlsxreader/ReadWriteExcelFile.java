package accenture.javaxlsxreader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadWriteExcelFile {

       @SuppressWarnings("deprecation")
       public static void readXLSFile() throws IOException
       {
              InputStream ExcelFileToRead = new FileInputStream("C:/testeXLX/Test.xls");
              HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);

              HSSFSheet sheet=wb.getSheetAt(0);
              HSSFRow row; 
              HSSFCell cell;

              Iterator rows = sheet.rowIterator();

              while (rows.hasNext())
              {
                     row=(HSSFRow) rows.next();
                     Iterator cells = row.cellIterator();
                     
                     while (cells.hasNext())
                     {
                           cell=(HSSFCell) cells.next();
              
                           if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING)
                           {
                                  System.out.print(cell.getStringCellValue()+" ");
                           }
                           else if(cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC)
                           {
                                  System.out.print(cell.getNumericCellValue()+" ");
                           }
                           else
                           {
                                  //U Can Handel Boolean, Formula, Errors
                           }
                     }
                     System.out.println();
              }
       
       }
       
       public static void writeXLSFile() throws IOException {
              
              String excelFileName = "C:/testeXLX/Test.xls";//name of excel file

              String sheetName = "Sheet1";//name of sheet

              HSSFWorkbook wb = new HSSFWorkbook();
              HSSFSheet sheet = wb.createSheet(sheetName) ;

              //iterating r number of rows
              for (int r=0;r < 1500; r++ )
              {
                     HSSFRow row = sheet.createRow(r);
                     System.out.println("Linha: "+r);
                     //iterating c number of columns
                     for (int c=0;c < 30; c++ )
                     {
                           HSSFCell cell = row.createCell(c);
                           
                           cell.setCellValue("Cell "+r+" "+c);
                     }
              }
              
              FileOutputStream fileOut = new FileOutputStream(excelFileName);
              
              //write this workbook to an Outputstream.
              wb.write(fileOut);
              fileOut.flush();
              fileOut.close();
       }
       
       public static void readXLSXFile() throws IOException
       {
    	   		System.out.println("opa01");
              InputStream ExcelFileToRead = new FileInputStream("C:/testeXLX/Test120mil.xlsx");
              System.out.println("opa02");
              XSSFWorkbook  wb = new XSSFWorkbook(ExcelFileToRead);
              System.out.println("opa03");
              XSSFWorkbook test = new XSSFWorkbook(); 
              System.out.println("opa04");
              XSSFSheet sheet = wb.getSheetAt(0);
              System.out.println("opa05");
              XSSFRow row; 
              System.out.println("opa06");
              XSSFCell cell;
              System.out.println("opa07");
              int linhanum = 0;
              
              Iterator rows = sheet.rowIterator();
              System.out.println("opa08");
              while (rows.hasNext())
              {
            	  System.out.println("opa09");
                     row=(XSSFRow) rows.next();
                     System.out.println("opa10");
                     Iterator cells = row.cellIterator();
                     
                     System.out.println("Linha: "+linhanum++);
                     
                     while (cells.hasNext())
                     {
                           cell=(XSSFCell) cells.next();
              
                           if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING)
                           {
                                  System.out.print(cell.getStringCellValue()+" ");
                           }
                           else if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC)
                           {
                                  System.out.print(cell.getNumericCellValue()+" ");
                           }
                           else
                           {
                                  //U Can Handel Boolean, Formula, Errors
                           }
                     }
                     System.out.println();
              }
       
       }
       
       public static void writeXLSXFile() throws IOException {
              
              String excelFileName = "C:/testeXLX/Test.xlsx";//name of excel file

              String sheetName = "Sheet1";//name of sheet

              XSSFWorkbook wb = new XSSFWorkbook();
              XSSFSheet sheet = wb.createSheet(sheetName) ;

              //iterating r number of rows
              for (int r=0;r < 100000; r++ )
              {
                     System.out.println("Linha: "+r);
                     XSSFRow row = sheet.createRow(r);

                     //iterating c number of columns
                     for (int c=0;c < 50; c++ )
                     {
                           XSSFCell cell = row.createCell(c);
       
                           cell.setCellValue("Cell "+r+" "+c);
                     }
              }

              FileOutputStream fileOut = new FileOutputStream(excelFileName);

              //write this workbook to an Outputstream.
              wb.write(fileOut);
              fileOut.flush();
              fileOut.close();
       }

       public static void main(String[] args) throws IOException {
              System.out.println("opa1");
              //writeXLSFile();
              //readXLSFile();
              System.out.println("opa2");
              //writeXLSXFile();
              readXLSXFile();
              System.out.println("opa3");
       }

}
