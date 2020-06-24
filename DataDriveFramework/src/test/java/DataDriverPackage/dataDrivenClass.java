package DataDriverPackage;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDrivenClass {

	public ArrayList<String> getData(String testcaseName) throws IOException {
		FileInputStream fis = new FileInputStream(
				"/Users/surajkute/eclipse-workspace/DataDriveFramework/ExcelFile/ExcelSheet.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int sheets = workbook.getNumberOfSheets();

		ArrayList<String> a = new ArrayList<String>();
		for (int i = 0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("MainSheet")) {
				XSSFSheet sheet = workbook.getSheetAt(i);

				// Step1: Identify TestCases column by scanning the entire 1st row
				Iterator<Row> rows = sheet.iterator();
				Row firstrow = rows.next();
				Iterator<Cell> ce = firstrow.cellIterator();
				int k = 0;
				int column = 0;

				while (ce.hasNext()) {
					Cell value = ce.next();
					if (value.getStringCellValue().equalsIgnoreCase("Testcases")) {
						// desired column
						column = k;

					}
					k++;
				}
				System.out.println(column);

				// once column is identified then scan entire testcase column to identify
				// purchase testcase row

				while (rows.hasNext()) {
					Row r = rows.next();
					if (r.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName)) {
						// after we grab purchase testcase row=pull all the data of that row into test

						Iterator<Cell> cv = r.cellIterator();
						while (cv.hasNext()) {
							// System.out.println(cv.next().getStringCellValue());
							a.add(cv.next().getStringCellValue());
						}
					}

				}

			}
		}

		return a;
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

	}

}
