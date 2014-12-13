/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.obleasnoob.controller;

import com.obleasnoob.entity.Sales;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author idea
 */
public class ExcelGenerator {
    
    HSSFWorkbook workbook;
    
    public File getFile(String fileName ){
        File file = null;
        try{
            file = new File(fileName);
            if(!file.exists()) {
               createEmptyFile(fileName);
            }
        }catch(IOException ex){
            Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex);
            try {
                createEmptyFile(fileName);
            } catch (IOException ex1) {
                Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex1);
            }
        }
        return file;
    }
    
    public void createEmptyFile(String fileName) throws FileNotFoundException, IOException{
        FileOutputStream fileOut =  new FileOutputStream(fileName);
        workbook = new HSSFWorkbook();
        HSSFSheet sheet =  workbook.createSheet("FirstSheet");  
        workbook.write(fileOut);
        fileOut.close();
    }
    
    public void writeSalesInExcelFile(String fileName, int indexSheet, Sales sale){
        FileInputStream file = null;
        FileOutputStream outFile = null;
        try {
            file = new FileInputStream(getFile(fileName));
            
            HSSFWorkbook workbook = new HSSFWorkbook(file);
            HSSFSheet sheet = workbook.getSheetAt(indexSheet);
            writeSalesSheet(sheet, sale);
            file.close();
            
            FileOutputStream out = new FileOutputStream(new File(fileName));
            workbook.write(out);
            out.close();
            System.out.println("Excel written successfully..");
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex);
        }catch(IOException ex){
            Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                file.close();
            } catch (IOException ex) {
                Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
    
    public void writeSalesSheet(HSSFSheet sheet, Sales sale){
        int indexRow = getRowTarget();
        int indexCol = 0;
        
        Row row = sheet.createRow(indexRow);
        createCell(row, indexCol, sale.getSaleDate());
        indexCol++;
        createCell(row, indexCol, sale.getOperatingRevenues());
        indexCol++;
        createCell(row, indexCol, sale.getCostSales());
        indexCol++;
        createCell(row, indexCol, sale.getGrossProfit());
        indexCol++;
        createCell(row, indexCol, sale.getOperatingExpenditure());
        indexCol++;
        createCell(row, indexCol, sale.getUAII());
        indexCol++;
        createCell(row, indexCol, sale.getNonOperatingIncome());
        indexCol++;
        createCell(row, indexCol, sale.getNonOperatingExpenses());
        indexCol++;
        createCell(row, indexCol, sale.getNetIncomeDaily());
        indexCol++;
        for(Double hour:sale.getHoursPerStations()){
            createCell(row, indexCol, hour);
            indexCol++;
        }
        createCell(row, indexCol, sale.getTotalHours());
        indexCol++;
        createCell(row, indexCol, sale.getEarningsPerHour());
        indexCol++;
    }
    
    public void createCell(Row row, int indexCol, Object obj){
        Cell cell = row.createCell(indexCol);
        if(obj instanceof Date) {
//            CellStyle cellStyle = workbook.createCellStyle();
//            cellStyle.setDataFormat(workbook.createDataFormat().getFormat("m/d/yy h:mm"));
            cell.setCellValue((Date)obj);
//            cell.setCellStyle(cellStyle);
        }else if(obj instanceof Boolean){
            cell.setCellValue((Boolean)obj);
        }else if(obj instanceof String){
            cell.setCellValue((String)obj);
        }else if(obj instanceof Double){
            cell.setCellValue((Double)obj);
        }
    }
    
    public int getRowTarget(){
        return 1;
    }

    public HSSFWorkbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(HSSFWorkbook workbook) {
        this.workbook = workbook;
    }
    
}
