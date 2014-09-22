import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

public class Main {

	public static void main (String[] args) throws IOException  {
		
		File file = new File("/home/chau/Documents/samples/test.html");
		file.createNewFile();
		FileWriter writer = new FileWriter(file);
		
		XLSReader xlsReader = new XLSReader("/home/chau/Documents/samples/Book1-1.xls");
		List<String> pages = xlsReader.getHTMLPages();
		
		writer.write(pages.get(0));
		writer.flush();
		writer.close();
		
		xlsReader.close();
		
	}
	
}
