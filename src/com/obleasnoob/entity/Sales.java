/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.obleasnoob.entity;

import java.util.Date;

/**
 *
 * @author idea
 */
public class Sales {
    private Date saleDate;
    private Double operatingRevenues;
    private Double costSales;
    private Double grossProfit;
    private Double operatingExpenditure;
    private Double UAII;
    private Double nonOperatingIncome;
    private Double nonOperatingExpenses;
    private Double netIncomeDaily;
    private Double[] hoursPerStations;
    private Double totalHours;
    private Double earningsPerHour;

    public Sales() {
        hoursPerStations = new Double[5];
    }

    public Date getSaleDate() {
        return saleDate;
    }

    public void setSaleDate(Date saleDate) {
        this.saleDate = saleDate;
    }

    public Double getOperatingRevenues() {
        return operatingRevenues;
    }

    public void setOperatingRevenues(Double operatingRevenues) {
        this.operatingRevenues = operatingRevenues;
    }

    public Double getCostSales() {
        return costSales;
    }

    public void setCostSales(Double costSales) {
        this.costSales = costSales;
    }

    public Double getGrossProfit() {
        return grossProfit;
    }

    public void setGrossProfit(Double grossProfit) {
        this.grossProfit = grossProfit;
    }

    public Double getOperatingExpenditure() {
        return operatingExpenditure;
    }

    public void setOperatingExpenditure(Double operatingExpenditure) {
        this.operatingExpenditure = operatingExpenditure;
    }

    public Double getUAII() {
        return UAII;
    }

    public void setUAII(Double UAII) {
        this.UAII = UAII;
    }

    public Double getNonOperatingIncome() {
        return nonOperatingIncome;
    }

    public void setNonOperatingIncome(Double nonOperatingIncome) {
        this.nonOperatingIncome = nonOperatingIncome;
    }

    public Double getNonOperatingExpenses() {
        return nonOperatingExpenses;
    }

    public void setNonOperatingExpenses(Double nonOperatingExpenses) {
        this.nonOperatingExpenses = nonOperatingExpenses;
    }

    public Double getNetIncomeDaily() {
        return netIncomeDaily;
    }

    public void setNetIncomeDaily(Double netIncomeDaily) {
        this.netIncomeDaily = netIncomeDaily;
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

    public Double getEarningsPerHour() {
        return earningsPerHour;
    }

    public void setEarningsPerHour(Double earningsPerHour) {
        this.earningsPerHour = earningsPerHour;
    }
    
}

