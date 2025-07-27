import static com.kms.katalon.core.testdata.TestDataFactory.findTestData

import com.kms.katalon.core.configuration.RunConfiguration

import excelActivity.ExcelUtils


String user = findTestData('Data Files/LoginDataSheet').getValue('UserName', 1)

println(user)

String filePath = RunConfiguration.getProjectDir()+File.separator+'Data Files'+File.separator+'LoginData.xlsx'

ExcelUtils.writeToCell(filePath, 'Sheet1', 'Name', 1, user)

def fetchdata = ExcelUtils.readData_from_DataFile('LoginDataSheet', 'Subject', 2)

println('Subject is : '+fetchdata)

String fetch2 = ExcelUtils.readData_from_Excel(filePath, 'Sheet1', 'Marks', 3)

println('Marks is : '+fetch2)