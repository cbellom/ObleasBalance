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
    
    public int getProppertiesTargetRowValue() throws IOException {
        Properties prop = new Properties();
        String propFileName = "config.properties";

        InputStream inputStream = getClass().getClassLoader().getResourceAsStream(propFileName);

        if (inputStream != null) {
            prop.load(inputStream);
        } else {
            throw new FileNotFoundException("property file '" + propFileName + "' not found in the classpath");
        }
        
        String row = prop.getProperty("targetRow");
        return Integer.parseInt(row);
    }
    
    public void setPropertiesTargetRowValue(String value) throws IOException {
        Properties prop = new Properties();
        String propFileName = "config.properties";
        
        FileInputStream in = new FileInputStream("src/"+propFileName);
        Properties props = new Properties();
        props.load(in);
        in.close();

        FileOutputStream out = new FileOutputStream("src/"+propFileName);
        props.setProperty("targetRow", value);
        props.store(out, null);
        out.close();
    }
}
