/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.obleasnoob.controller;

import com.obleasnoob.entity.Inventory;
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
                    createEmptyExcelFile(fileName);                    
            }
        }catch(IOException ex){
            Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex);
            try {
                createEmptyExcelFile(fileName);
            } catch (IOException ex1) {
                Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex1);
            }
        }
        return file;
    }
    
    public void createEmptyExcelFile(String fileName) throws FileNotFoundException, IOException{
        
        String extention = getExtentionFile(fileName);
        FileOutputStream fileOut =  new FileOutputStream(fileName);
        if("xls".equals(extention)){
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet =  workbook.createSheet(propertiesController.getSheetName());
            workbook.write(fileOut);
        } else if("xlsx".equals(extention)){
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet =  workbook.createSheet(propertiesController.getSheetName());
            workbook.write(fileOut);
        }
        fileOut.close();
            
    }
    
    public void writeInXLSFile(String fileName, Object obj){
        FileInputStream file = null;
        FileOutputStream outFile = null;
        try {
            file = new FileInputStream(getFile(fileName));
            
            HSSFWorkbook workbook = new HSSFWorkbook(file);
            HSSFSheet sheet = workbook.getSheetAt(propertiesController.getSheetIndex());
            setWorkbookXLS(workbook);
            
            if(obj instanceof Inventory){
                writeInventorySheet(sheet, (Inventory)obj);                
            }else if(obj instanceof Sales){
                writeSalesSheet(sheet, (Sales)obj);
            }
            
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
    
    public void writeInXLSXFile(String fileName, Object obj){
        FileInputStream file = null;
        FileOutputStream outFile = null;
        try {
            file = new FileInputStream(getFile(fileName));
            
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(propertiesController.getSheetIndex());
            setWorkbookXLSX(workbook);
            
            if(obj instanceof Inventory){
                writeInventorySheet(sheet, (Inventory)obj);                
            }else if(obj instanceof Sales){
                writeSalesSheet(sheet, (Sales)obj);
            }
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
    
    public void writeSalesSheet(HSSFSheet sheet, Sales sale) throws IOException{
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
        
        indexRow++;
        propertiesController.setTargetRow(indexRow);
    }
    
    public void writeSalesSheet(XSSFSheet sheet, Sales sale) throws IOException{
        int indexRow = getRowTarget();
        int indexCol = 0;
        
        XSSFRow row = sheet.createRow(indexRow);
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
        
        indexRow++;
        propertiesController.setTargetRow(indexRow);
    }
    
    public void writeInventorySheet(HSSFSheet sheet, Inventory inventory) throws IOException{
        int indexRow = getRowTarget();
        int indexCol = getColTarget();
        
        Row row = sheet.createRow(indexRow);
        for(Double hour:inventory.getSalesPerStations()){
            createCell(row, indexCol, hour);
            indexCol++;
        }        
        for(Double hour:inventory.getHoursPerStations()){
            createCell(row, indexCol, hour);
            indexCol++;
        }        
        createCell(row, indexCol, inventory.getTotalHours());
        indexCol++;
        createCell(row, indexCol, inventory.getWafers());
        indexCol++;
        createCell(row, indexCol, inventory.getArequipe());
        indexCol++;
        createCell(row, indexCol, inventory.getCheese());
        indexCol++;
        createCell(row, indexCol, inventory.getBlackberry());
        indexCol++;
        createCell(row, indexCol, inventory.getCream());
        indexCol++;
        createCell(row, indexCol, inventory.getAzucar());
        indexCol++;
        createCell(row, indexCol, inventory.getRice());
        indexCol++;
        createCell(row, indexCol, inventory.getMilk());
        indexCol++;        
        createCell(row, indexCol, inventory.getGrapes());
        indexCol++;
        createCell(row, indexCol, inventory.getCinnamon());
        indexCol++;
        createCell(row, indexCol, inventory.getNail());
        indexCol++;
        createCell(row, indexCol, inventory.getFuel());
        indexCol++;
        createCell(row, indexCol, inventory.getNapkins());
        indexCol++;
        createCell(row, indexCol, inventory.getGlasses());
        indexCol++;
        createCell(row, indexCol, inventory.getGloves());
        indexCol++;
        createCell(row, indexCol, inventory.getTea());
        indexCol++;
        createCell(row, indexCol, inventory.getOthers());
        indexCol++;
        createCell(row, indexCol, inventory.getComments());
        indexCol++;
    }
    
    public void writeInventorySheet(XSSFSheet sheet, Inventory inventory) throws IOException{
        int indexRow = getRowTarget();
        int indexCol = getColTarget();
        
        Row row = sheet.createRow(indexRow);
        for(Double hour:inventory.getSalesPerStations()){
            createCell(row, indexCol, hour);
            indexCol++;
        }        
        for(Double hour:inventory.getHoursPerStations()){
            createCell(row, indexCol, hour);
            indexCol++;
        }        
        createCell(row, indexCol, inventory.getTotalHours());
        indexCol++;
        createCell(row, indexCol, inventory.getWafers());
        indexCol++;
        createCell(row, indexCol, inventory.getArequipe());
        indexCol++;
        createCell(row, indexCol, inventory.getCheese());
        indexCol++;
        createCell(row, indexCol, inventory.getBlackberry());
        indexCol++;
        createCell(row, indexCol, inventory.getCream());
        indexCol++;
        createCell(row, indexCol, inventory.getAzucar());
        indexCol++;
        createCell(row, indexCol, inventory.getRice());
        indexCol++;
        createCell(row, indexCol, inventory.getMilk());
        indexCol++;        
        createCell(row, indexCol, inventory.getGrapes());
        indexCol++;
        createCell(row, indexCol, inventory.getCinnamon());
        indexCol++;
        createCell(row, indexCol, inventory.getNail());
        indexCol++;
        createCell(row, indexCol, inventory.getFuel());
        indexCol++;
        createCell(row, indexCol, inventory.getNapkins());
        indexCol++;
        createCell(row, indexCol, inventory.getGlasses());
        indexCol++;
        createCell(row, indexCol, inventory.getGloves());
        indexCol++;
        createCell(row, indexCol, inventory.getTea());
        indexCol++;
        createCell(row, indexCol, inventory.getOthers());
        indexCol++;
        createCell(row, indexCol, inventory.getComments());
        indexCol++;
    }
    
    public void createCell(Row row, int indexCol, Object obj){
        Cell cell = row.createCell(indexCol);
        if(obj instanceof Date) {
                Date date = (Date)obj;            
                cell.setCellValue(date);
        }else if(obj instanceof Boolean){
            cell.setCellValue((Boolean)obj);
            cell.setCellType(Cell.CELL_TYPE_BOOLEAN);
        }else if(obj instanceof String){
            cell.setCellValue((String)obj);
            cell.setCellType(Cell.CELL_TYPE_STRING);
        }else if(obj instanceof Double){
            cell.setCellValue((Double)obj);
            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
        }
    }
    
    public int getRowTarget(){
        return propertiesController.getTargetRow();
    }
    
    public int getColTarget(){
        return propertiesController.getTargetCol();
    }
  
    public String getExtentionFile(String fileName){
        String extention;
        int indexPoint = fileName.lastIndexOf(".");
        extention = fileName.substring(indexPoint+1, fileName.length());
        return extention;
    }
    
    public void writeSales(Sales sale){
        try {
            propertiesController.getPropertiesValues();
            
            String fileName = propertiesController.getFileName();        
            String extention = getExtentionFile(fileName);
            
            if("xls".equals(extention))
                writeInXLSFile(fileName, sale);
            else if("xlsx".equals(extention))
                writeInXLSXFile(fileName, sale);
        
            propertiesController.updatePropertiesValues();
        } catch (IOException ex) {
            Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public HSSFWorkbook getWorkbookXLS() {
        return workbookXLS;
    }

    public void setWorkbookXLS(HSSFWorkbook workbook) {
        this.workbookXLS = workbook;
    }

    public XSSFWorkbook getWorkbookXLSX() {
        return workbookXLSX;
    }

    public void setWorkbookXLSX(XSSFWorkbook workbookXLSX) {
        this.workbookXLSX = workbookXLSX;
    }
    
}
