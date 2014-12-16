/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.obleasnoob.controller;
import java.util.Date;

/**
 *
 * @author idea
 */
public class Validator {
    
    public boolean isNumberField(String input){
        try {
            Integer number = Integer.parseInt(input);
            return true;
        } catch (Exception e) {
            return false;
        }
    }
    
    public boolean isDateField(String input){
        try {
            Date date = new Date(input);
            return true;
        } catch (Exception e) {
            return false;
        }
    }
}
