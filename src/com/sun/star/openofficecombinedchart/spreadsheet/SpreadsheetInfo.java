/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */



package com.sun.star.openofficecombinedchart.spreadsheet;

/**
 *
 * @author othman
 */
public class SpreadsheetInfo {
    private String[] columnHeader;
    private int      max_row;
    private int      min_row;
    private String   rangeString;
    private String   sheetName;

    public SpreadsheetInfo(String[] columnHeader, String rangeName) {
        this.columnHeader = columnHeader;
        this.rangeString  = rangeName;
        this.parseRangeString();
    }

    public void parseRangeString() {
        String[] sheet = rangeString.split("\\$");

        if ((sheet == null) || (sheet.length < 6)) {
            return;
        }

        sheet[1]       = sheet[1].replace('.', ' ');
        sheet[3]       = sheet[3].replace(':', ' ');
        this.sheetName = sheet[1].trim();
        this.min_row   = Integer.parseInt(sheet[3].trim());
        this.max_row   = Integer.parseInt(sheet[5].trim());
    }

    public String[] getColumnHeader() {
        return columnHeader;
    }

    public void setColumnHeader(String[] columnHeader) {
        this.columnHeader = columnHeader;
    }

    public String getRangeString() {
        return rangeString;
    }

    public int getMax_row() {
        return max_row;
    }

    public int getMin_row() {
        return min_row;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setRangeString(String sheetName) {
        this.rangeString = sheetName;
    }
}
