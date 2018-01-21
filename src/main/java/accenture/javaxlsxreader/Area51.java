package accenture.javaxlsxreader;

import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;

public class Area51 {

	public static void main(String[] args) {
        
        LargeExcelXLSXReader reader = new LargeExcelXLSXReader("C:/testeXLX/Test.xlsx");
        
        try {
        	TimeMeasurer cron = new TimeMeasurer(); //start timer
			
        	reader.load();
        	
			cron.restart(); //check time
			
			String[] lines = reader.iterate(1); 
			
			cron.restart(); //check time
			
			//for (String line : lines) {
			//	System.out.println(line);
			//}
			
			cron.stop();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
}