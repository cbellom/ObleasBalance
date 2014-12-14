/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.obleasnoob.entity;

/**
 *
 * @author idea
 */
public class Inventory {    
    private Double[] salesPerStations;
    private Double[] hoursPerStations;
    private Double totalHours;
    private Double wafers;
    private Double arequipe;
    private Double cheese;
    private Double blackberry;
    private Double cream;
    private Double Azucar;
    private Double rice;
    private Double Milk;
    private Double grapes;
    private Double cinnamon;
    private Double nail;
    private Double fuel;
    private Double napkins;
    private Double glasses;
    private Double gloves;
    private Double Tea;
    private Double others;
    private String comments;

    public Inventory() {
        salesPerStations = new Double[5];
        hoursPerStations = new Double[5];
    }

    public Double[] getSalesPerStations() {
        return salesPerStations;
    }

    public void setSalesPerStations(Double[] salesPerStations) {
        this.salesPerStations = salesPerStations;
    }

    public Double[] getHoursPerStations() {
        return hoursPerStations;
    }

    public void setHoursPerStations(Double[] hoursPerStations) {
        this.hoursPerStations = hoursPerStations;
    }

    public Double getTotalHours() {
        return totalHours;
    }

    public void setTotalHours(Double totalHours) {
        this.totalHours = totalHours;
    }

    public Double getWafers() {
        return wafers;
    }

    public void setWafers(Double wafers) {
        this.wafers = wafers;
    }

    public Double getArequipe() {
        return arequipe;
    }

    public void setArequipe(Double arequipe) {
        this.arequipe = arequipe;
    }

    public Double getCheese() {
        return cheese;
    }

    public void setCheese(Double cheese) {
        this.cheese = cheese;
    }

    public Double getBlackberry() {
        return blackberry;
    }

    public void setBlackberry(Double blackberry) {
        this.blackberry = blackberry;
    }

    public Double getCream() {
        return cream;
    }

    public void setCream(Double cream) {
        this.cream = cream;
    }

    public Double getAzucar() {
        return Azucar;
    }

    public void setAzucar(Double Azucar) {
        this.Azucar = Azucar;
    }

    public Double getRice() {
        return rice;
    }

    public void setRice(Double rice) {
        this.rice = rice;
    }

    public Double getMilk() {
        return Milk;
    }

    public void setMilk(Double Milk) {
        this.Milk = Milk;
    }

    public Double getGrapes() {
        return grapes;
    }

    public void setGrapes(Double grapes) {
        this.grapes = grapes;
    }

    public Double getCinnamon() {
        return cinnamon;
    }

    public void setCinnamon(Double cinnamon) {
        this.cinnamon = cinnamon;
    }

    public Double getNail() {
        return nail;
    }

    public void setNail(Double nail) {
        this.nail = nail;
    }

    public Double getFuel() {
        return fuel;
    }

    public void setFuel(Double fuel) {
        this.fuel = fuel;
    }

    public Double getNapkins() {
        return napkins;
    }

    public void setNapkins(Double napkins) {
        this.napkins = napkins;
    }

    public Double getGlasses() {
        return glasses;
    }

    public void setGlasses(Double glasses) {
        this.glasses = glasses;
    }

    public Double getGloves() {
        return gloves;
    }

    public void setGloves(Double gloves) {
        this.gloves = gloves;
    }

    public Double getTea() {
        return Tea;
    }

    public void setTea(Double Tea) {
        this.Tea = Tea;
    }

    public Double getOthers() {
        return others;
    }

    public void setOthers(Double others) {
        this.others = others;
    }

    public String getComments() {
        return comments;
    }

    public void setComments(String comments) {
        this.comments = comments;
    }
    
}
