package accenture.javaxlsxreader;

import java.io.FileNotFoundException;

public class Area51 {

	public static void main(String[] args) {
        
        LargeExcelXLSXReader reader = new LargeExcelXLSXReader("C:\\testeXLX\\Carga 210_RI.xlsx");
        
        try {
        	TimeMeasurer cron = new TimeMeasurer(); //start timer
			
        	reader.load();
        	
			cron.restart(); //check time
			
        	reader.iterate(0); 
			
			cron.stop(); //check time
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
}