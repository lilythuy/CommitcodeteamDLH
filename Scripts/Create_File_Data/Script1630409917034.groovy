import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import java.io.FileInputStream as FileInputStream
import java.io.FileNotFoundException as FileNotFoundException
import java.io.IOException as IOException
import java.util.Date as Date
import org.apache.poi.xssf.usermodel.XSSFCell as XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow as XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet as XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook as XSSFWorkbook
import java.lang.String as String
import org.apache.poi.hssf.usermodel.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;


String valueName = Code

// Creating an instance of HSSFWorkbook.
HSSFWorkbook workbook = new HSSFWorkbook();

// Create two sheets in the excel document and name it First Sheet and
// Second Sheet.
HSSFSheet firstSheet = workbook.createSheet("FIRST SHEET");
HSSFSheet secondSheet = workbook.createSheet("SECOND SHEET");

// Manipulate the firs sheet by creating an HSSFRow which represent a
// single row in excel sheet, the first row started from 0 index. After
// the row is created we create a HSSFCell in this first cell of the row
// and set the cell value with an instance of HSSFRichTextString
// containing the words FIRST SHEET.

HSSFRow rowA = firstSheet.createRow(0);

HSSFCell cell_1 = rowA.createCell(0);
cell_1.setCellValue(new HSSFRichTextString("PLATE ID"));

HSSFCell cell_2 = rowA.createCell(1);
cell_2.setCellValue(new HSSFRichTextString("No"));

HSSFCell cell_3 = rowA.createCell(2);
cell_3.setCellValue(new HSSFRichTextString("Plate posision"));

HSSFCell cell_4 = rowA.createCell(3);
cell_4.setCellValue(new HSSFRichTextString("Barcode"));

HSSFCell cell_5 = rowA.createCell(4);
cell_5.setCellValue(new HSSFRichTextString("Gender"));

HSSFCell cell_6 = rowA.createCell(5);
cell_6.setCellValue(new HSSFRichTextString("Conc."));

HSSFCell cell_7 = rowA.createCell(6);
cell_7.setCellValue(new HSSFRichTextString("Unit"));

HSSFCell cell_8 = rowA.createCell(7);
cell_8.setCellValue(new HSSFRichTextString("ul aliquote to send 4ug"));

HSSFCell cell_9 = rowA.createCell(8);
cell_9.setCellValue(new HSSFRichTextString("A260"));

HSSFCell cell_10 = rowA.createCell(9);
cell_10.setCellValue(new HSSFRichTextString("A280"));

HSSFCell cell_11 = rowA.createCell(10);
cell_11.setCellValue(new HSSFRichTextString("260/280"));

HSSFCell cell_12 = rowA.createCell(11);
cell_12.setCellValue(new HSSFRichTextString("Sample Type"));

HSSFCell cell_13 = rowA.createCell(12);
cell_13.setCellValue(new HSSFRichTextString("Factor"));

HSSFCell cell_14 = rowA.createCell(13);
cell_14.setCellValue(new HSSFRichTextString("Results"));

HSSFCell cell_15 = rowA.createCell(14);
cell_15.setCellValue(new HSSFRichTextString("DNA form"));

//--- Row 2 Data ---- //

HSSFRow rowB = firstSheet.createRow(1);

HSSFCell cell_B1 = rowB.createCell(0);
cell_B1.setCellValue(new HSSFRichTextString("PLATE_" + valueName));

HSSFCell cell_B2 = rowB.createCell(1);
cell_B2.setCellValue(new HSSFRichTextString("1"));

HSSFCell cell_B3 = rowB.createCell(2);
cell_B3.setCellValue(new HSSFRichTextString("A1"));

HSSFCell cell_B4 = rowB.createCell(3);
cell_B4.setCellValue(new HSSFRichTextString(valueName));

HSSFCell cell_B5 = rowB.createCell(4);
cell_B5.setCellValue(new HSSFRichTextString(""));

HSSFCell cell_B6 = rowB.createCell(5);
cell_B6.setCellValue(new HSSFRichTextString("8.5"));

HSSFCell cell_B7 = rowB.createCell(6);
cell_B7.setCellValue(new HSSFRichTextString("ng/µl"));


// To write out the workbook into a file we need to create an output
// stream where the workbook content will be written to.
String filePath = "D:\\TESTER\\Katalon\\batchLab_PLATE_" + valueName + ".xls"
FileOutputStream outFile = new FileOutputStream(new File(filePath))
workbook.write(outFile)