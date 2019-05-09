import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SplitFile {

	private final String fileName;
	private final int maxRows;
	private final String path;
	private final String userfilename = "";
	public static int filecount;
	public static String taskname;
	public static int rowcounter;
	private SXSSFRow headRow;
	private int rcountflag;
	private boolean mdir;
	private String newFileName;

	public SplitFile(String fileName, final int maxRows, String filepath, String userfilename)
			throws FileNotFoundException {
		path = filepath;
		taskname = userfilename;
		this.fileName = fileName;
		this.maxRows = maxRows;
		System.out.println("Iniciando a aplicacao");

		File file = new File(fileName);
		FileInputStream inputStream = new FileInputStream(file);
		try {
			/* Leitura do arquivo Excel. */
			// OPCPackage pkg = OPCPackage.open(new File(fileName));
			Workbook workbook = new XSSFWorkbook(inputStream);
			XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(0);
			System.out.println("Selecionando Aba");

			/* Only split if there are more rows than the desired amount. */
			if (sheet.getPhysicalNumberOfRows() >= maxRows) {
				List<SXSSFWorkbook> wbs = splitWorkbook(workbook);
				writeWorkBooks(wbs);
			}

		} catch (EncryptedDocumentException | IOException e) {
			e.printStackTrace();
		}
	}

	private List<SXSSFWorkbook> splitWorkbook(Workbook workbook) {

		List<SXSSFWorkbook> workbooks = new ArrayList<SXSSFWorkbook>();

		SXSSFWorkbook wb = new SXSSFWorkbook();
		SXSSFSheet sh = (SXSSFSheet) wb.createSheet();

		SXSSFRow newRow;
		setHeadRow(null);
		SXSSFCell newCell;
		String headCellarr[] = new String[50];

		int rowCount = 0;
		int colCount = 0;
		int headflag = 0;
		setRcountflag(0);
		int cols = 0;

		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(0);

		// sheet.createFreezePane(0, 1);
		int i = 0;
		rowcounter++;
		for (Row row : sheet) {
			if (i == 0) {

				/* Time to create a new workbook? */
				int j = 0;
				for (Cell cell : row) {

					// newCell = headRow.createCell(colCount++);
					headCellarr[j] = cell.toString();
					j++;
				}
				cols = j;
				colCount = 0;
				i++;
			} else {
				break;
			}

		}

		for (Row row : sheet) {

			/* Time to create a new workbook? */
			if (rowCount == maxRows) {
				headflag = 1;
				workbooks.add(wb);
				wb = new SXSSFWorkbook();
				sh = (SXSSFSheet) wb.createSheet();
				rowCount = 0;

			}
			if (headflag == 1) {
				newRow = (SXSSFRow) sh.createRow(rowCount++);
				headflag = 0;
				for (int k = 0; k < cols; k++) {
					newCell = (SXSSFCell) newRow.createCell(colCount++);
					newCell.setCellValue(headCellarr[k]);

				}
				colCount = 0;
				newRow = (SXSSFRow) sh.createRow(rowCount++);

				for (Cell cell : row) {
					newCell = (SXSSFCell) newRow.createCell(colCount++);
					if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
						newCell.setCellValue("-");
					} else {
						newCell = setValue(newCell, cell);
					}
				}
				colCount = 0;

			} else {
				rowcounter++;
				newRow = (SXSSFRow) sh.createRow(rowCount++);
				for (int cn = 0; cn < row.getLastCellNum(); cn++) {

					// If the cell is missing from the file, generate a blank one
					// (Works by specifying a MissingCellPolicy)
					Cell cell = row.getCell(cn, Row.CREATE_NULL_AS_BLANK);

					// Print the cell for debugging
					// System.out.println("CELL: " + cn + " --> " + cell.toString());
					newCell = (SXSSFCell) newRow.createCell(cn);
					if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
						newCell.setCellValue(cell.getNumericCellValue());
					} else {
						newCell.setCellValue(cell.toString());
					}
				}
			}
		}

		/* Only add the last workbook if it has content */
		if (wb.getSheetAt(0).getPhysicalNumberOfRows() > 0) {
			workbooks.add(wb);
		}
		return workbooks;
	}

	/*
	 * Grabbing cell contents can be tricky. We first need to determine what type of
	 * cell it is.
	 */
	private SXSSFCell setValue(SXSSFCell newCell, Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			// newCell.setCellValue(cell.getRichStringCellValue().getString());
			newCell.setCellValue(cell.toString());
			break;
		case Cell.CELL_TYPE_NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				newCell.setCellValue(cell.getDateCellValue());
			} else {
				newCell.setCellValue(cell.getNumericCellValue());
				// newCell.setCellValue(cell.toString());
			}
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			newCell.setCellValue(cell.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_FORMULA:
			newCell.setCellFormula(cell.getCellFormula());
			break;
		case Cell.CELL_TYPE_BLANK:
			newCell.setCellValue("");
			break;
		default:
			System.out.println("Nao foi possivel determinar tipo da celula");
			newCell.setCellValue(cell.toString());

		}
		return newCell;
	}

	/* Write all the workbooks to disk. */
	private void writeWorkBooks(List<SXSSFWorkbook> wbs) {
		FileOutputStream out;
		setMdir(new File(path).mkdir());

		try {
			for (int i = 0; i < wbs.size(); i++) {
				setNewFileName(fileName.substring(0, fileName.length() - 5));
				// out = new FileOutputStream(new File(newFileName + "_" + (i + 1) + ".xlsx"));
				out = new FileOutputStream(new File(path + "/" + taskname + "_" + (i + 1) + ".xlsx"));
				wbs.get(i).write(out);
				out.close();
				System.out.println("Escrevendo..." + i);
				filecount++;
			}
			System.out.println(userfilename);
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public int sendtotalrows() {
		return rowcounter;
	}

	public static void main(String[] args) throws FileNotFoundException {
		// This will create a new workbook every 1000 rows.
		// new Splitter(filename.xlsx, No of split rows, filepath, newfilename);
		new SplitFile("C:\\Teste\\RCB.xlsx", 30000, "c:\\teste", "RCB"); // No of rows to split: 10 K
	}

	public SXSSFRow getHeadRow() {
		return headRow;
	}

	public void setHeadRow(SXSSFRow headRow) {
		this.headRow = headRow;
	}

	public int getRcountflag() {
		return rcountflag;
	}

	public void setRcountflag(int rcountflag) {
		this.rcountflag = rcountflag;
	}

	public boolean isMdir() {
		return mdir;
	}

	public void setMdir(boolean mdir) {
		this.mdir = mdir;
	}

	public String getNewFileName() {
		return newFileName;
	}

	public void setNewFileName(String newFileName) {
		this.newFileName = newFileName;
	}
}
