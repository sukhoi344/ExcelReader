package office.reader.excel;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import office.reader.util.ColorUtil;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * @author chau.thai
 */
public class XLSSheetReader {
	
	private static final float WIDTH_RATIO = 1.0f;
	
	private HSSFWorkbook workBook;
	private HSSFSheet sheet;
	
	private StringBuilder sb;
	
	/** Column span means cell span horizontally.  RowIndex -> Span[ ] */
	private Map<Integer, List<Span>> spanColsMap;
	
	/** Row span means cell span vertically. ColIndex -> Span[ ] */
	private Map<Integer, List<Span>> spanRowsMap;
	
	private int lastRowIndex = 0;
	private int lastColIndex = 0;
	
	public XLSSheetReader(HSSFWorkbook workBook, HSSFSheet sheet) {
		this.sheet = sheet;
		this.workBook = workBook;

		sb = new StringBuilder().append(XLSReader.HTML_START);
		spanColsMap = new HashMap<>();
		spanRowsMap = new HashMap<>(); 
		
		lastColIndex = getLastColIndex();
		lastRowIndex = sheet.getLastRowNum();
	}
	
	public String getHTML() {
		setupSpans();
		addStyle();
		addBody();
		
		return sb.append(XLSReader.HTML_END).toString();
	}
	
	private void addBody() {
		sb.append("<body><table>");
		
		// Setup column width
		sb.append("<col width=\"10\">");
		for(int i = 0; i <= lastColIndex; i++)
			sb.append("<col width=\"" + (sheet.getColumnWidth(i) * WIDTH_RATIO) + "\">");
		
		// Add header row
		addHeaderRow();
		
		// Add rows
		addRows();
		
		sb.append("</table></body>");
	}
	
	private void addRows() {

		int rowCount = 1;

		// Iterate through rows
		Iterator<Row> rowIter = sheet.rowIterator();

		while(rowIter.hasNext()) {
			boolean decrementBlankNum = false;
			
			HSSFRow row = (HSSFRow) rowIter.next();
			Iterator<Cell> cellIter = row.cellIterator();
			
			// Check if needed to add blank rows before the current row
			addBlankRows(rowCount, row.getRowNum() - rowCount + 1);
			rowCount += row.getRowNum() - rowCount + 1;
				
			// Add row index
			sb.append("<tr><td class = \"header\">  " + (rowCount) + "  </td>");
			
			int lastColIndex = -1;	// last column index of the last cell

			// Iterate through cells
			while(cellIter.hasNext()) {
				Cell cell = cellIter.next();
				String cellContent = cell.toString(); 
				
				// Get cell style
				StringBuilder style = getStyleString(cell);
	
				int colIndex = cell.getColumnIndex();
				int blankCellsNum = colIndex - lastColIndex - 1;	// Number of blank cells to be added before the current cell
				lastColIndex = colIndex;
				
				// Check for special case, which decrements blank cells by 1
				if(decrementBlankNum) {
					blankCellsNum--;
					decrementBlankNum = false;
				}
				
				// Add blank cells before the current cell
				addBlankCells(blankCellsNum);
				
				// Check for merged cells
				List<Span> colSpans = spanColsMap.get(cell.getRowIndex());
				List<Span> rowSpans = spanRowsMap.get(cell.getColumnIndex());
				
				int colSpan = 0;
				int rowSpan = 0;
				
				// Check horizontally
				if(colSpans != null) {
					for(Span span : colSpans) {
						if(span.firstIndex == colIndex) {
							colSpan = span.spanRange;
							//blankCellsNum -= colSpan ;
							int spanCount = colSpan - 1;
							
							// Skip adding cells because of the merged cell
							for(int i = 0; i < spanCount; i++) {
								if(cellIter.hasNext()) {
									cellIter.next();
									lastColIndex++;
								}
							}
							
							break;	// Got the desired blank cell numbers, stop the loop
						}
					}
				}
				
				// Check vertically
				if(rowSpans != null) {
					for(Span span : rowSpans) {
						if(span.firstIndex == cell.getRowIndex()) {
							rowSpan = span.spanRange;
							break;
						}
						
						if(cell.getRowIndex() > span.firstIndex && cell.getRowIndex() <= span.lastIndex) {
							colSpan = 0;
							decrementBlankNum = true;
							break;
						}
					}
				}
				
				sb.append("<td " + style 
							+ ((colSpan > 0) ? "colspan=\"" + colSpan + "\"" : "") 
							+ ((rowSpan > 0) ? "rowspan=\"" + rowSpan + "\"" : "")
							+ ">"
							+ cellContent + "</td>");
			}
			
			// Add blank cells to the end of the last cell
			addBlankCells(this.lastColIndex - lastColIndex);
			
			// End row
			sb.append("</tr>");
			rowCount++;
		}
	}

	private StringBuilder getStyleString(Cell cell) {
		StringBuilder sb = new StringBuilder();
		sb.append("style=\"");
		
		String cellContent = cell.toString();
		CellStyle cellStyle = cell.getCellStyle();
		
		try {

			// Get background color
			Short fillBackgroundIndex = cellStyle.getFillBackgroundColor();
			String colorHex = ColorUtil.getColorString(workBook, fillBackgroundIndex);
			if(colorHex.equals("#000000"))
				colorHex = "#FFFFFF";
			
			// Get font style
			HSSFFont font = workBook.getFontAt(cellStyle.getFontIndex());

			short fontColorIndex = font.getColor();
			short fontBoldWeight = font.getBoldweight();
			boolean italic = font.getItalic();
			short heightInPoints = font.getFontHeightInPoints();
			
			String fontColorHex = ColorUtil.getColorString(workBook, fontColorIndex);
			
			// Get border style
			String borderBottomStyle = getBorderStyle(cellStyle.getBorderBottom());
			String borderTopStyle = getBorderStyle(cellStyle.getBorderTop());
			String borderLeftStyle = getBorderStyle(cellStyle.getBorderLeft());
			String borderRightStyle = getBorderStyle(cellStyle.getBorderRight());
			
			String borderBottomColor = ColorUtil.getColorString(workBook, cellStyle.getBottomBorderColor());
			String borderTopColor = ColorUtil.getColorString(workBook, cellStyle.getTopBorderColor());
			String borderLeftColor = ColorUtil.getColorString(workBook, cellStyle.getLeftBorderColor());
			String borderRightColor = ColorUtil.getColorString(workBook, cellStyle.getRightBorderColor());
			
			// Add styles
			if(!borderLeftStyle.isEmpty())
				sb.append("border-left:" + borderLeftStyle + " " + borderLeftColor + ";");
			if(!borderRightStyle.isEmpty())
				sb.append("border-right:" + borderRightStyle + " " + borderRightColor + ";");
			if(!borderTopStyle.isEmpty())
				sb.append("border-top:" + borderTopStyle + " " + borderTopColor + ";");
			if(!borderBottomStyle.isEmpty())
				sb.append("border-bottom:" + borderBottomStyle + " " + borderBottomColor + ";");
			
			if(italic)	
				sb.append("font-style:italic;");
			sb.append("font-weight:" + fontBoldWeight + ";");
			sb.append("font-size:" + heightInPoints + "0%;");
			sb.append("color:" + fontColorHex + ";");
			sb.append("background:" + colorHex + ";");

		} catch (Exception e) {
			return new StringBuilder("");
		}
		
		return sb.append("\"");
	}
	
	private void addHeaderRow() {
		StringBuilder headerBuilder = new StringBuilder();
		
		headerBuilder.append("<tr class = \"header\">");
		headerBuilder.append("<td> </td>");
		
		for(int i = 0; i <= lastColIndex; i++) 
			headerBuilder.append("<td>" +indexToAlphabet(i) + "</td>");
		
		sb.append(headerBuilder);
	}

	private void addStyle() {
		sb.append("<head><style>");
		
		sb.append("table, th, td {"
				+ 	"border: 1px solid #C4C4C4;"
				+	"border-collapse: collapse;"
				+ "}");
		
		sb.append("th, td {"  
				+	   "padding: 5px;"  
				+	   "text-align: center" 
				+ "}");
		
		sb.append("tr.header {"
				+	"background: #EEEEEE;"
				+ "}");
		
		sb.append("td.header {"
				+	"background: #EEEEEE;"
				+ "}");
		
		sb.append("</style>");
	}

	private void setupSpans() {
		int mergeNum = sheet.getNumMergedRegions();
		
		for(int i = 0; i < mergeNum; i++) {
			CellRangeAddress cellRangeAddress = sheet.getMergedRegion(i);
			
			int firstCol = cellRangeAddress.getFirstColumn();
			int lastCol = cellRangeAddress.getLastColumn();
			int firstRow = cellRangeAddress.getFirstRow();
			int lastRow = cellRangeAddress.getLastRow();
			
			// Add span rows
			for(int col = firstCol; col <= lastCol; col++) {
				Span span = new Span(firstRow, lastRow);
				
				if(span.spanRange > 1) {
					if(spanRowsMap.get(col) == null) 
						spanRowsMap.put(col, new ArrayList<Span>());
					
					spanRowsMap.get(col).add(span);
				}
			}
			
			// Add span columns
			for(int row = firstRow; row <= lastRow; row++) {
				Span span = new Span(firstCol, lastCol);
				
				if(span.spanRange > 1) {
					if(spanColsMap.get(row) == null)
						spanColsMap.put(row, new ArrayList<Span>());
					
					spanColsMap.get(row).add(span);
				}
			}
		}
	}
	
	private int getLastColIndex() {
		int lastColIndex = 1;
		
		Iterator<Row> rowIter = sheet.rowIterator();
		while(rowIter.hasNext()) {
			Row row = rowIter.next();
			int lastCelNum = row.getLastCellNum();
			
			if(lastCelNum > lastColIndex)
				lastColIndex = lastCelNum;
		}
		
		return lastColIndex - 1;
	}
	
	private String getBorderStyle(short styleIndex) {
		switch (styleIndex) {
		case CellStyle.BORDER_DASHED:
			return "dashed 2px";
		case CellStyle.BORDER_DOTTED:
			return "dotted 2px";
		case CellStyle.BORDER_DOUBLE:
			return "double";
		case CellStyle.BORDER_NONE:
			return "";
		case CellStyle.BORDER_THIN:
		case CellStyle.BORDER_MEDIUM:
			return "solid 2px";
		case CellStyle.BORDER_THICK:
			return "solid";
		case CellStyle.BORDER_DASH_DOT:
		case CellStyle.BORDER_DASH_DOT_DOT:
		case CellStyle.BORDER_MEDIUM_DASH_DOT:
		case CellStyle.BORDER_MEDIUM_DASH_DOT_DOT:
		case CellStyle.BORDER_MEDIUM_DASHED:
		case CellStyle.BORDER_SLANTED_DASH_DOT:
			return "solid 2px";
			
		default:
			return "";	
		}
	}
	
	/** Add blank cells into the row */
	private void addBlankCells(int n) {
		for(int i = 0; i < n; i++)
			sb.append("<td></td>");
	}
	
	/**
	 *  Add rows of full blank cells into the sheet
	 * @param currentRowIndex starting index of rows to be added
	 * @param n number of rows to be added
	 */
	private void addBlankRows(int currentRowIndex, int n) {
		if(currentRowIndex > 0 && n > 0) {
			for(int i = 0; i < n; i++) {
				sb.append("<tr><td class = \"header\">  " + (currentRowIndex++) + "  </td>");
				addBlankCells(lastColIndex + 1);
				sb.append("</tr>");
			}
		}
	}
	
	/**
	 *  Index begins at 0. Alphabets from (A - ZZ)
	 * @param index
	 * @return
	 */
	private static String indexToAlphabet(int index) {
		if(index < 0)
			return "";
		
		if(index < 26) 
			return Character.toString((char) ('A' + index));
		
		int firstDigit = ((index / 26) % 26) - 1 + 65;
		firstDigit = (firstDigit > 64) ? firstDigit : 90;
		
		int secondDigit = index % 26 + 65;
		
		String firstChar = Character.toString((char) secondDigit);
		String secondChar = Character.toString((char) firstDigit);
		
		return secondChar + firstChar;
	}
	
	class Span {
		int firstIndex;
		int lastIndex;
		int spanRange;
		
		public Span(int firstIndex, int lastIndex) {
			this.firstIndex = firstIndex;
			this.lastIndex = lastIndex;
			
			spanRange = lastIndex - firstIndex + 1;
		}
	}
	
	class SpanRange {
		
		int rowIndex;
		int colIndex;
		
		int totalSpanRows = 0;
		int totalSpanCols = 0;
		
		public SpanRange(int rowIndex, int colIndex) {
			this.rowIndex = rowIndex;
			this.colIndex = colIndex;
		}
		
		public int getTotalSpanRows() {
			int total = 0;
			List<Span> rowSpans = spanRowsMap.get(colIndex);
			
			for(Span rowSpan : rowSpans) 
				if(rowSpan.lastIndex < rowIndex) 
					total += rowSpan.spanRange;
			
			return total;
		}
		
		public int getTotalSpanCols() {
			int total = 0;
			List<Span> colSpans = spanColsMap.get(rowIndex);
			
			for(Span colSpan : colSpans) 
				if(colSpan.lastIndex < colIndex) 
					total += colSpan.spanRange;
			
			return total;
		}
	}
}
