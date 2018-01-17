package accenture.javaxlsxreader;

import java.io.FileNotFoundException;

public class Area51 {

	public static void main(String[] args) {
        
        LargeExcelXLSXReader reader = new LargeExcelXLSXReader("C:\\testeXLX\\Test600mil.xlsx");
        try {
			reader.load();
			reader.iterate();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
}