package excelActivity

import static com.kms.katalon.core.testdata.TestDataFactory.findTestData

import org.apache.poi.xssf.usermodel.*

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testdata.TestDataFactory

class ExcelUtils {

	@Keyword
	static void writeToCell(String filePath, String SheetName, String columnHeader, int rowNumber, String dataValue) {
		FileInputStream fis = new FileInputStream(filePath)
		XSSFWorkbook workbook = new XSSFWorkbook(fis)
		fis.close()

		XSSFSheet sheet = workbook.getSheet(SheetName)
		XSSFRow headerRow = sheet.getRow(0)

		int columnIndex = -1
		for (int i = 0; i < headerRow.getLastCellNum(); i++) {
			XSSFCell headerCell = headerRow.getCell(i)
			if (headerCell != null && headerCell.toString().trim().equalsIgnoreCase(columnHeader.trim())) {
				columnIndex = i
				break
			}
		}

		if (columnIndex == -1) {
			throw new IllegalArgumentException("Column '" + columnHeader + "' not found.")
		}

		XSSFRow row = sheet.getRow(rowNumber)
		if (row == null) {
			row = sheet.createRow(rowNumber)
		}

		XSSFCell cell = row.getCell(columnIndex)
		if (cell == null) {
			cell = row.createCell(columnIndex)
		}

		cell.setCellValue(dataValue)

		FileOutputStream fos = new FileOutputStream(filePath)
		workbook.write(fos)
		fos.close()
		workbook.close()
	}
	
	@Keyword
    static String readData_from_Excel(String filePath, String sheetName, String columnHeader, int rowNumber) {
        FileInputStream fis = new FileInputStream(filePath)
        XSSFWorkbook workbook = new XSSFWorkbook(fis)
        XSSFSheet sheet = workbook.getSheet(sheetName)

        XSSFRow headerRow = sheet.getRow(0)
        XSSFRow dataRow = sheet.getRow(rowNumber)

        int colIndex = -1
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            if (headerRow.getCell(i).toString().trim() == columnHeader) {
                colIndex = i
                break
            }
        }

        String value = dataRow.getCell(colIndex).toString()

        workbook.close()
        fis.close()

        return value
    }
	
	@Keyword
	static String readData_from_DataFile(String DataFileName, String columnName, int rowNumber) {
		def data = TestDataFactory.findTestData(DataFileName)
		String value = data.getValue(columnName, rowNumber)
		return value
	}
}
