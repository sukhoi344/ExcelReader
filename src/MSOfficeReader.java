import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * @author chau.thai
 */
public abstract class MSOfficeReader {
	
	protected static final String HTML_START = "<!DOCTYPE html><html>";
	protected static final String HTML_END = "</html>";
	
	private InputStream is;
	
	public MSOfficeReader(File file) throws IOException {
		this(new FileInputStream(file));
	}
	
	public MSOfficeReader(String filePath) throws IOException {
		this(new File(filePath));
	}
	
	public MSOfficeReader(InputStream is) throws IOException {
		this.is = is;
		onCreatePOIDocument(new POIFSFileSystem(is));
	}
	
	/**
	 * Instantiate POIDocument derived class object here. 
	 * This method is called at the end of the constructor
	 * 
	 * @param fileSystem used to create POIDocument object
	 */
	protected abstract void onCreatePOIDocument(POIFSFileSystem fileSystem) throws IOException;
	
	/**
	 * Get list of the file content in HTML format. The list
	 * can contains pages of Doc, or Excel, depends on the derived
	 * class's implementation.
	 * 
	 * @return list of HTML pages
	 */
	public abstract List<String> getHTMLPages();
	
	/**
	 * Clear all the resources
	 */
	public void close() {
		if(is != null) {
			try {
				is.close();
			} catch (Exception e) {}
		}
	}
}
