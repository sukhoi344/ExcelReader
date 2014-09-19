import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;


public class XLSReader extends MSOfficeReader {

	
	private HSSFWorkbook workBook;
	private int totalSheets;
	
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
		totalSheets = workBook.getNumberOfSheets();
	}

	@Override
	public List<String> getHTMLPages() {
		
		List<String> listHTML = new ArrayList<String>();
		
		for(int index = 0; index < totalSheets; index++) {
			HSSFSheet sheet = workBook.getSheetAt(index);
			XLSSheetReader sheetReader = new XLSSheetReader(sheet);
			
			listHTML.add(sheetReader.getHTML());
		}
		
		return listHTML;
	}
}
