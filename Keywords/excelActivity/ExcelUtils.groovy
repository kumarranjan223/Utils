package excelActivity

import org.apache.poi.xssf.usermodel.*

import com.kms.katalon.core.annotation.Keyword

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
	
}
