/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.obleasnoob.controller;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.Properties;

/**
 *
 * @author idea
 */
public class PropertiesController {
    
    private String fileName;
    private String sheetName;
    private int sheetIndex;
    private int targetCol;
    private int targetColDate;
    private String propFileName;

    public PropertiesController() {
        fileName = "Balance.xls";
        sheetName = "Invent. y sueldos";
        sheetIndex = 0;
        targetCol = 0;
        targetColDate = 0;
        propFileName = "config.properties";
    }
        
    public void getPropertiesValues() throws IOException {
        Properties prop = new Properties();
        
        FileInputStream inputStream = new FileInputStream(propFileName);
        Properties props = new Properties();
        
        if (inputStream != null) {
            prop.load(inputStream);
        } else {
            throw new FileNotFoundException("property file '" + propFileName + "' not found in the classpath");
        }
        
        fileName = prop.getProperty("fileName");
        sheetName = prop.getProperty("sheetName");
        sheetIndex = Integer.parseInt(prop.getProperty("sheetIndex"));
        targetCol = Integer.parseInt(prop.getProperty("targetCol"));
        targetColDate = Integer.parseInt(prop.getProperty("targetColDate"));
        inputStream.close();
    }
        
    public void updatePropertiesValues() throws IOException {
        Properties prop = new Properties();
        
        FileInputStream inputFile = new FileInputStream(propFileName);
        Properties props = new Properties();
        props.load(inputFile);
        inputFile.close();

        FileOutputStream out = new FileOutputStream(propFileName);
        props.setProperty("fileName", fileName);
        props.setProperty("sheetName", sheetName);
        props.setProperty("sheetIndex", sheetIndex+"");
        props.setProperty("targetCol", targetCol+"");
        props.setProperty("targetColDate", targetColDate+"");
        props.store(out, null);
        out.close();
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public int getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public int getTargetCol() {
        return targetCol;
    }

    public void setTargetCol(int targetCol) {
        this.targetCol = targetCol;
    }

    public int getTargetColDate() {
        return targetColDate;
    }

    public void setTargetColDate(int targetColDate) {
        this.targetColDate = targetColDate;
    }
    
}
