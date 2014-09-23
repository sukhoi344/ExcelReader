package office.reader.word;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import office.reader.MSOfficeReader;

/**
 * @author chau.thai
 */
public class DocReader extends MSOfficeReader {

	private HWPFDocument document;
	
	public DocReader(InputStream is) throws IOException {
		super(is);
	}
	
	public DocReader(File file) throws IOException {
		super(file);
	}
	
	public DocReader(String filePath) throws IOException {
		super(filePath);
	}

	@Override
	protected void onCreatePOIDocument(POIFSFileSystem fileSystem)
			throws IOException {
		document = new HWPFDocument(fileSystem);
	}

	@Override
	public List<String> getHTMLPages() {
		WordExtractor extractor = new WordExtractor(document);
		List<String> list = new ArrayList<>();
		
		StringBuilder sb = new StringBuilder();
		sb.append(HTML_START);
		sb.append("<head></head><body>");
		
		
		String[] paragraphs = extractor.getParagraphText();
		for(String para : paragraphs) {
			sb.append(para);
		}
		
		try {
			extractor.close();
		} catch (IOException e) {}
		
		sb.append("</body>");
		sb.append(HTML_END);
		
		list.add(sb.toString());
		
		return list;
	}
	
	
}
