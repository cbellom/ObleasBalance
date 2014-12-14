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
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author idea
 */
public class ExcelGenerator {
    
    HSSFWorkbook workbookXLS;
    XSSFWorkbook workbookXLSX;
    PropertiesController propertiesController = new PropertiesController();
    
    public File getFile(String fileName ){
        File file = null;
        try{
            file = new File(fileName);
            if(!file.exists()) {
               createEmptyXLSFile(fileName);
            }
        }catch(IOException ex){
            Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex);
            try {
                createEmptyXLSFile(fileName);
            } catch (IOException ex1) {
                Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex1);
            }
        }
        return file;
    }
    
    public void createEmptyXLSFile(String fileName) throws FileNotFoundException, IOException{
        FileOutputStream fileOut =  new FileOutputStream(fileName);
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet =  workbook.createSheet("FirstSheet");
        workbook.write(fileOut);
        fileOut.close();
    }
    
    public void writeSalesInXLSFile(String fileName, int indexSheet, Sales sale){
        FileInputStream file = null;
        FileOutputStream outFile = null;
        try {
            file = new FileInputStream(getFile(fileName));
            
            HSSFWorkbook workbook = new HSSFWorkbook(file);
            HSSFSheet sheet = workbook.getSheetAt(indexSheet);
            setWorkbookXLS(workbook);
            writeSalesSheetXLS(sheet, sale);
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
    
    public void writeSalesSheetXLS(HSSFSheet sheet, Sales sale) throws IOException{
        int indexRow = getRowTarget();
        int indexCol = 0;
        
        Row row = sheet.createRow(indexRow);
        createCellXLS(row, indexCol, sale.getSaleDate());
        indexCol++;
        createCellXLS(row, indexCol, sale.getOperatingRevenues());
        indexCol++;
        createCellXLS(row, indexCol, sale.getCostSales());
        indexCol++;
        createCellXLS(row, indexCol, sale.getGrossProfit());
        indexCol++;
        createCellXLS(row, indexCol, sale.getOperatingExpenditure());
        indexCol++;
        createCellXLS(row, indexCol, sale.getUAII());
        indexCol++;
        createCellXLS(row, indexCol, sale.getNonOperatingIncome());
        indexCol++;
        createCellXLS(row, indexCol, sale.getNonOperatingExpenses());
        indexCol++;
        createCellXLS(row, indexCol, sale.getNetIncomeDaily());
        indexCol++;
        for(Double hour:sale.getHoursPerStations()){
            createCellXLS(row, indexCol, hour);
            indexCol++;
        }
        createCellXLS(row, indexCol, sale.getTotalHours());
        indexCol++;
        createCellXLS(row, indexCol, sale.getEarningsPerHour());
        indexCol++;
        
        indexRow++;
        propertiesController.setPropertiesTargetRowValue(indexRow+"");
    }
    
    public void createCellXLS(Row row, int indexCol, Object obj){
        Cell cell = row.createCell(indexCol);
        if(obj instanceof Date) {
                Date date = (Date)obj;            
                cell.setCellValue(date);
                HSSFCellStyle dateCellStyle = getWorkbookXLS().createCellStyle();
                short df = getWorkbookXLS().createDataFormat().getFormat("dd-MMM-YYYY");
                dateCellStyle.setDataFormat(df);
                cell.setCellStyle(dateCellStyle);
        }else if(obj instanceof Boolean){
            cell.setCellValue((Boolean)obj);
            cell.setCellType(HSSFCell.CELL_TYPE_BOOLEAN);
        }else if(obj instanceof String){
            cell.setCellValue((String)obj);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
        }else if(obj instanceof Double){
            cell.setCellValue((Double)obj);
            cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
        }
    }
    
    public int getRowTarget(){
        int row = 0;
        try {
            row = propertiesController.getPropertiesTargetRowValue();
        } catch (IOException ex) {
            Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex);
        }
        return row;
    }
    
    public int getColTarget(){
        int row = 0;
        try {
            row = propertiesController.getPropertiesTargetRowValue();
        } catch (IOException ex) {
            Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex);
        }
        return row;
    }

    public HSSFWorkbook getWorkbookXLS() {
        return workbookXLS;
    }

    public void setWorkbookXLS(HSSFWorkbook workbook) {
        this.workbookXLS = workbook;
    }
    
    public String getExtentionFile(String fileName){
        String extention;
        int indexPoint = fileName.lastIndexOf(".");
        extention = fileName.substring(indexPoint+1, fileName.length());
        return extention;
    }
    
}
