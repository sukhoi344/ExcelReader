package office.reader.excel;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import office.reader.MSOfficeReader;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;


public class XLSReader extends MSOfficeReader {
	
	private HSSFWorkbook workBook;
	private int totalSheets;
	private List<HSSFSheet> sheets;
	
	public XLSReader(InputStream is) throws IOException {
		super(is);
	}
	
	public XLSReader(File file) throws IOException {
		super(file);
	}
	
	public XLSReader(String filePath) throws IOException {
		super(filePath);
	}
	
	@Override
	protected void onCreatePOIDocument(POIFSFileSystem fileSystem) throws IOException {
		workBook = new HSSFWorkbook(fileSystem);
		sheets = new ArrayList<HSSFSheet>();
		totalSheets = workBook.getNumberOfSheets();
		
		// Get sheets
		for(int index = 0; index < totalSheets; index++) 
			sheets.add(workBook.getSheetAt(index));
	}

	@Override
	public List<String> getHTMLPages() {
		List<String> listHTML = new ArrayList<String>();
		
		for(HSSFSheet sheet : sheets) {
			XLSSheetReader sheetReader = new XLSSheetReader(workBook, sheet);
			listHTML.add(sheetReader.getHTML());
		}
		
		return listHTML;
	}
	
	/** Get list of all the sheet names */
	public List<String> getSheetNames() {
		List<String> names = new ArrayList<String>();
		
		for(HSSFSheet sheet : sheets) 
			names.add(sheet.getSheetName());
		
		return  names;
	}
	
	/**
	 * Get sheet name at index. 
	 * @param index
	 * @return Sheet name. null String if index is invalid
	 */
	public String getSheetNameAtIndex(int index) {
		if(workBook == null || index < 0 || index >= totalSheets)
			return null;
		
		return workBook.getSheetName(index);
	}
}
