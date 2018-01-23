package accenture.javaxlsxreader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import com.monitorjbl.xlsx.StreamingReader;


/**
 * Based on https://github.com/monitorjbl/excel-streaming-reader
 */
public class LargeExcelXLSXReader {
	
	private int numRowsInMemory = 100; // number of rows to keep in memory (defaults to 10)
	private int bufferSize = 4096; // buffer size to use when reading InputStream to file (defaults to 1024)
	private String filepath = null;
	private File inputFile = null;
	private Workbook workbook = null;
	private InputStream is = null; // InputStream or File for XLSX file (required)
	private ArrayList<String[]> arrayInfo = new ArrayList<String[]>();
	
	//private static final java.util.logging.Logger LOGGER = java.util.logging.Logger.getLogger(LargeExcelXLSXReader.class.getName());
	
	/**
	 * Construtor padrão
	 */
	public LargeExcelXLSXReader() {
		
	}
	
	/**
	 * Construtor com argumentos
	 * @param filepath
	 */
	public LargeExcelXLSXReader(String filepath) {
		setFilepath(filepath);
	}

	/**
	 * Carrega o arquivo XLSX na memória (parcialmente)
	 * @throws FileNotFoundException
	 */
	public void load() throws FileNotFoundException {
		setInputFile(new File(getFilepath()));
		setIs(new FileInputStream(getInputFile()));
		setWorkbook(
				StreamingReader.builder()
				.rowCacheSize(getNumRowsInMemory())
				.bufferSize(getBufferSize())
				.open(getIs())
		);
	}
	
	/**
	 * Carrega o arquivo XLSX na memória (parcialmente)
	 * 
	 * @param filepath
	 * @throws FileNotFoundException
	 */
	public void load(String filepath) throws FileNotFoundException {
		setFilepath(filepath);
		this.load();
	}
	
	/**
	 * Itera sobre todas as linhas do arquivo excel.
	 * 
	 * @param debugMode
	 * @return
	 */
	public List<String[]> iterate(int debugMode) {
		
		if(getWorkbook() == null)
			return null;
		    	
		int linha = 1;	//debug
		int planilha = 1;	//capturar apenas a primeira planilha
		
		if(debugMode!=0) {
			System.out.println("Iniciando leitura do arquivo ...");			
		}
		
    	for (Sheet sheet : getWorkbook()){
   			if(planilha>1) { // capturar apenas a primeira planilha
				continue;
			}
    		
    		if(debugMode==2) {
    			System.out.println(sheet.getSheetName());
    		}
    		
    		for (Row r : sheet) {
 
    			if(debugMode==2) {
    				System.out.print("linha: " + (linha++) + " [");    				
    			}
    			
    			String[] row = new String[r.getLastCellNum()];
    			int cellNum = 0;
    			for (Cell c : r) {
    				if(debugMode==2) {
    					System.out.print(c.getStringCellValue() + ";");
    				}
    				
    				row[cellNum] = c.getStringCellValue();
    				cellNum++;
    			}
    			
    			arrayInfo.add(row);
    			
    			if(debugMode==2) {
    				System.out.println("]");
    			}
    		}
    		
    		planilha++;
		}
    	
    	System.out.println("Leitura finalizada com sucesso.");
		return arrayInfo;
	}
	
	public int getBufferSize() {
		return bufferSize;
	}

	public void setBufferSize(int bufferSize) {
		this.bufferSize = bufferSize;
	}

	public int getNumRowsInMemory() {
		return numRowsInMemory;
	}

	public void setNumRowsInMemory(int numRowsInMemory) {
		this.numRowsInMemory = numRowsInMemory;
	}

	public Workbook getWorkbook() {
		return workbook;
	}

	public void setWorkbook(Workbook workbook) {
		this.workbook = workbook;
	}

	public InputStream getIs() {
		return is;
	}

	public void setIs(InputStream is) {
		this.is = is;
	}

	public String getFilepath() {
		return filepath;
	}

	public void setFilepath(String filepath) {
		this.filepath = filepath;
	}

	public File getInputFile() {
		return inputFile;
	}

	public void setInputFile(File inputFile) {
		this.inputFile = inputFile;
	}	
}