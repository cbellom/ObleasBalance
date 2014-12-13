/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package obleasnoob;

import com.obleasnoob.controller.ExcelGenerator;
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

/**
 *
 * @author idea
 */
public class ObleasNoob {
    private static ExcelGenerator excelGenerator = new ExcelGenerator();

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        Sales sales = new Sales();
        sales.setSaleDate(new Date());
        sales.setCostSales(Double.parseDouble("2.1"));
        excelGenerator.writeSalesInExcelFile("Balance.xls", 0, sales);
    }
    
}
