/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package obleasnoob;

import com.obleasnoob.controller.ExcelGenerator;
import com.obleasnoob.controller.PropertiesController;
import com.obleasnoob.entity.Inventory;
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
        Inventory inv = new Inventory();
        Date x = new Date();
        x.setYear(2990);
        inv.setDate(x);
        inv.setBlackberry(Double.parseDouble("1.1"));
        inv.setArequipe(Double.parseDouble("2.0"));
        inv.setSugar(Double.parseDouble("4.0"));
        excelGenerator.writeData(inv);
        
//        Sales sales = new Sales();
//        sales.setSaleDate(new Date());
//        sales.setCostSales(Double.parseDouble("2.1"));
//        excelGenerator.writeData(sales);
    }    
}
