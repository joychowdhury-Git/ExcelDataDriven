package ExcelDrivenTest.ExcelDataDriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDrivenTest {

	public ArrayList<String> getdata(String testcasename) throws Exception
	{
		ArrayList<String> a = new ArrayList<String>();
		FileInputStream fis = new FileInputStream("C:\\Users\\joych\\NewWorkspace\\ExcelDataDriven\\ExcelData.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		// Count the number of sheets present in excel

		int sheetcount = workbook.getNumberOfSheets();

		for (int i = 0; i < sheetcount; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("Sheet1")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				// identify the testcase column from 1st row

				Iterator<Row> rows = sheet.iterator(); // sheet is collection of rows
				Row firstrow = rows.next(); // indicate first row

				Iterator<Cell> cols = firstrow.cellIterator(); // indicate first column

				int k = 0;
				int Column = 0; // use for index
				while (cols.hasNext()) {
					Cell cellvalue = cols.next();
					if (cellvalue.getStringCellValue().equalsIgnoreCase("TestCase")) {
						Column = k;
					}

					k++;
				}
				// System.out.println(Column);
				// Scan the entire first column and identify the purchase testcase from 1st
				// column

				while (rows.hasNext()) {
					Row rowvalue = rows.next();

					if (rowvalue.getCell(Column).getStringCellValue().equalsIgnoreCase(testcasename)) {
						Iterator<Cell> purchasevaluecell = rowvalue.cellIterator();

						while (purchasevaluecell.hasNext()) {
							
							Cell Pcell = purchasevaluecell.next();
							
							if(Pcell.getCellType() == CellType.STRING) {
							
							a.add(Pcell.getStringCellValue());
						}
							else {
								
								a.add(NumberToTextConverter.toText(Pcell.getNumericCellValue()));
								
							}
						}

					}
						
				}
				
			}
		}
		return a;
	}

}
