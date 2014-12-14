/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package obleasnoob;

import com.obleasnoob.controller.ExcelGenerator;
import com.obleasnoob.controller.PropertiesController;
import com.obleasnoob.entity.Sales;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author idea
 */
public class ObleasNoob {
    private static ExcelGenerator excelGenerator = new ExcelGenerator();

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException {
        Sales sales = new Sales();
        sales.setSaleDate(new Date());
        sales.setCostSales(Double.parseDouble("2.1"));
        excelGenerator.writeSales("Balance.xlsx", 0, sales);
        
        String excelFileName = "Test.xlsx";//name of excel file
 
		String sheetName = "Sheet1";//name of sheet
 
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet(sheetName) ;
 
		//iterating r number of rows
		for (int r=0;r < 5; r++ )
		{
			XSSFRow row = sheet.createRow(r);
 
			//iterating c number of columns
			for (int c=0;c < 5; c++ )
			{
				XSSFCell cell = row.createCell(c);
	
				cell.setCellValue("Cell "+r+" "+c);
			}
		}
 
		FileOutputStream fileOut = new FileOutputStream(excelFileName);
 
		//write this workbook to an Outputstream.
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
    }
    
}
