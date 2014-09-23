import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

import office.reader.excel.XLSReader;

public class Main {

	public static void main (String[] args) throws IOException  {
		
		File file = new File("/Users/chauthai/Documents/android/MacateReader/test.html");
		file.createNewFile();
		FileWriter writer = new FileWriter(file);
		
		XLSReader xlsReader = new XLSReader("/Users/chauthai/Documents/android/MacateReader/Book1-1.xls");
		List<String> pages = xlsReader.getHTMLPages();
		
		writer.write(pages.get(0));
		writer.flush();
		writer.close();
		
		xlsReader.close();
		
	}
	
}
