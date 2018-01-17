package accenture.javaxlsxreader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSXtoCSVconverter {

    public static void xlsx(File inputFile, File outputFile) {
    	
        // For storing data into CSV files
        StringBuffer data = new StringBuffer();

        try {
        	System.out.println("print1");
            FileOutputStream fos = new FileOutputStream(outputFile);
            // Get the workbook object for XLSX file
            System.out.println("print2");
            XSSFWorkbook wBook = (XSSFWorkbook) WorkbookFactory.create(inputFile);
            //XSSFWorkbook wBook = new XSSFWorkbook(new FileInputStream(inputFile));
            // Get first sheet from the workbook
            System.out.println("print3");
            XSSFSheet sheet = wBook.getSheetAt(0);
            Row row;
            Cell cell;
            // Iterate through each rows from first sheet
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {
                row = rowIterator.next();

                // For each row, iterate through each columns
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {

                    cell = cellIterator.next();

                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_BOOLEAN:
                            data.append(cell.getBooleanCellValue() + ";");

                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            data.append(cell.getNumericCellValue() + ";");

                            break;
                        case Cell.CELL_TYPE_STRING:
                            data.append(cell.getStringCellValue() + ";");
                            break;

                        case Cell.CELL_TYPE_BLANK:
                            data.append("" + ";");
                            break;
                        default:
                            data.append(cell + ";");

                    }
                    
                    
                }
                
                data.append("\r\n");
            }

            fos.write(data.toString().getBytes());
            fos.close();

        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
    }
    //testing the application 

    public static void main(String[] args) {
        //reading file from desktop
        File inputFile = new File("C:\\testeXLX\\Test10mil.xlsx");
        //writing excel data to csv 
        File outputFile = new File("C:\\testeXLX\\Test10milcsv.csv");
        xlsx(inputFile, outputFile);
    }
}