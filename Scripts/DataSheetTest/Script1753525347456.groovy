import static com.kms.katalon.core.testdata.TestDataFactory.findTestData

import com.kms.katalon.core.configuration.RunConfiguration

import excelActivity.ExcelUtils


String user = findTestData('Data Files/LoginDataSheet').getValue('UserName', 1)

println(user)

String filePath = RunConfiguration.getProjectDir()+File.separator+'Data Files'+File.separator+'LoginData.xlsx'

ExcelUtils.writeToCell(filePath, 'Sheet1', 'Sotel', 2, user)


