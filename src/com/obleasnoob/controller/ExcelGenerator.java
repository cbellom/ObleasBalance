/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.obleasnoob.controller;

import com.obleasnoob.entity.Inventory;
import com.obleasnoob.entity.Sales;
import java.awt.Frame;
import java.awt.HeadlessException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author idea
 */
public class ExcelGenerator {
    
    private int indexRowModified;
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
    
    public void writeInXLSFile(String fileName, Object obj, Frame frame){
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
            JOptionPane.showMessageDialog(frame, "Registro guardado exitosamente", "Guardado", JOptionPane.INFORMATION_MESSAGE);
        }catch(IOException | HeadlessException ex){
            JOptionPane.showMessageDialog(frame, ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                file.close();
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(frame, ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
    
    public void writeInXLSXFile(String fileName, Object obj, Frame frame){
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
            
            JOptionPane.showMessageDialog(frame, "Registro guardado exitosamente", "Guardado", JOptionPane.INFORMATION_MESSAGE);     
        }catch(IOException | HeadlessException ex){
            JOptionPane.showMessageDialog(frame, ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                file.close();
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(frame, ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
    
    public void writeSalesSheet(HSSFSheet sheet, Sales sale) throws IOException{
        
    }
    
    public void writeSalesSheet(XSSFSheet sheet, Sales sale) throws IOException{
        
    }
    
    public void writeInventorySheet(HSSFSheet sheet, Inventory inventory) throws IOException{
        int indexRow = getIndexRowByDate(inventory.getDate(), sheet, getColDateTarget());
        int indexCol = getColTarget();
        
        if(indexRow<0){
            indexRow = sheet.getLastRowNum();    
            Row row = sheet.createRow(indexRow);        
            createInventory(inventory, row, indexCol);
        }else{
            updateInventory(sheet, inventory, indexRow, indexCol);
        }  
    }
    
    public void writeInventorySheet(XSSFSheet sheet, Inventory inventory) throws IOException{
        int indexRow = getIndexRowByDate(inventory.getDate(), sheet, getColDateTarget());
        int indexCol = getColTarget();
        
        if(indexRow<0){
            indexRow = sheet.getLastRowNum();    
            Row row = sheet.createRow(indexRow);        
            createInventory(inventory, row, indexCol);
        }else{
            updateInventory(sheet, inventory, indexRow, indexCol);
        }
        indexRowModified = indexRow;
    }
    
    private void createInventory(Inventory inventory, Row row, int indexCol){
        createCell(row, getColDateTarget(), inventory.getDate());
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
        createCell(row, indexCol, inventory.getSugar());
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
    
    private void updateInventory(XSSFSheet sheet, Inventory inventory, int indexRow, int indexCol){        
        for(Double hour:inventory.getSalesPerStations()){
            updateCell(sheet, indexRow, indexCol, hour);
            indexCol++;
        }        
        for(Double hour:inventory.getHoursPerStations()){
            updateCell(sheet, indexRow, indexCol, hour);
            indexCol++;
        }        
        updateCell(sheet, indexRow, indexCol, inventory.getTotalHours());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getWafers());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getArequipe());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getCheese());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getBlackberry());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getCream());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getSugar());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getRice());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getMilk());
        indexCol++;        
        updateCell(sheet, indexRow, indexCol, inventory.getGrapes());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getCinnamon());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getNail());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getFuel());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getNapkins());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getGlasses());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getGloves());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getTea());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getOthers());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getComments());
        indexCol++;
    }
    
    private void updateInventory(HSSFSheet sheet, Inventory inventory, int indexRow, int indexCol){        
        for(Double hour:inventory.getSalesPerStations()){
            updateCell(sheet, indexRow, indexCol, hour);
            indexCol++;
        }        
        for(Double hour:inventory.getHoursPerStations()){
            updateCell(sheet, indexRow, indexCol, hour);
            indexCol++;
        }        
        updateCell(sheet, indexRow, indexCol, inventory.getTotalHours());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getWafers());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getArequipe());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getCheese());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getBlackberry());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getCream());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getSugar());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getRice());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getMilk());
        indexCol++;        
        updateCell(sheet, indexRow, indexCol, inventory.getGrapes());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getCinnamon());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getNail());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getFuel());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getNapkins());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getGlasses());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getGloves());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getTea());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getOthers());
        indexCol++;
        updateCell(sheet, indexRow, indexCol, inventory.getComments());
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
    
    public void updateCell(Sheet sheet, int indexRow, int indexCol, Object obj){        
        Cell cell;
        cell = sheet.getRow(indexRow).getCell(indexCol);
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
    
    public int getColDateTarget(){
        return propertiesController.getTargetColDate();
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
    
    public void writeData(Object obj, Frame frame){
        try {
            propertiesController.getPropertiesValues();
            
            String fileName = propertiesController.getFileName();        
            String extention = getExtentionFile(fileName);
            
            if("xls".equals(extention))
                writeInXLSFile(fileName, obj, frame);
            else if("xlsx".equals(extention))
                writeInXLSXFile(fileName, obj, frame);
            
        } catch (IOException ex) {
            Logger.getLogger(ExcelGenerator.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public int getIndexRowByDate(Date date, XSSFSheet sheet, int indexCellDate){
        Iterator<Row> rows=sheet.rowIterator();
        while (rows.hasNext()) {
            XSSFRow row = (XSSFRow) rows.next();
            Cell cellDate = row.getCell(indexCellDate);
            try{
                if(cellDate!=null && cellDate.getDateCellValue().getDate()==date.getDate()){
                    System.out.println("fecha:" + cellDate.getDateCellValue());
                    return row.getRowNum();
                }
            }catch(Exception ex){
                System.out.println("tipo de celda invalido");
            }
        }    
        return -1;
    }
    
    public int getIndexRowByDate(Date date, HSSFSheet sheet, int indexCellDate){
        Iterator<Row> rows=sheet.rowIterator();
        while (rows.hasNext()) {
            HSSFRow row = (HSSFRow) rows.next();
            Cell cellDate = row.getCell(indexCellDate);
            try{
                if(cellDate!=null && cellDate.getDateCellValue().getDate()==date.getDate()){
                    System.out.println("fecha:" + cellDate.getDateCellValue());
                    return row.getRowNum();
                }
            }catch(Exception ex){
                System.out.println("tipo de celda invalido");
            }
        }    
        return -1;
    }
    
    public int getLastIndexRow(XSSFSheet sheet){
        Iterator<Row> rows=sheet.rowIterator();
        int rowIndex=0;
        while (rows.hasNext()) {
            rows.next();
            rowIndex ++;
        }    
        return rowIndex;
    }
    
    public int getLastIndexRow(HSSFSheet sheet){
        Iterator<Row> rows=sheet.rowIterator();
        int rowIndex=0;
        while (rows.hasNext()) {
            rows.next();
            rowIndex ++;
        }    
        return rowIndex;
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
