package com.sioi.test_excel_in.jxl;

import java.util.ArrayList;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;

public class Cells {

    int MaxColumns = 0;
    int MaxRows = 0;

    private Cell[][] cells;

    /**
     * initialise Cells
     * @param sheet
     */
    Cells(Sheet sheet) {
        MaxColumns = sheet.getColumns();
        MaxRows = sheet.getRows();

        cells = new Cell[MaxRows + 1][MaxColumns + 1];

        for (int i = 0; i < MaxRows; i++) {
            for (int j = 0; j < MaxColumns; j++) {
                cells[i][j] = sheet.getCell(j, i);
            }

        }
    }

    /**
     * get String whole row by row index
     * @param row
     * @return List<String>
     */
    public List<String> getRow(int row) {

        List<String> rowsString = new ArrayList<String>();
        for (int i = 0; i < MaxColumns; i++) {
            rowsString.add(cells[row][i].getContents());
        }
        return rowsString;
    }

    /**
     * get String whole column by column index
     * @param column
     * @return List<String>
     */
    public List<String> getColumn(int column) {

        List<String> columnString = new ArrayList<String>();
        for (int i = 0; i < MaxRows; i++) {
            columnString.add(cells[i][column].getContents());
        }
        return columnString;
    }

    /**
     * get String by index row,column
     * @param row
     * @param column
     * @return String
     */
    public String getIndex(int row, int column) {

        return cells[row][column].getContents();
    }

    /**
     * print all table
     */
    public void getPrint() {

        for (int i = 0; i < MaxRows; i++) {

            for (int j = 0; j < MaxColumns; j++) {
                TablePrint(this.getIndex(i, j));
               // System.out.print(" ");
            }
            System.out.println();
        }
    }

    /**
     * print by 10s String
     * @param str
     */
    private static void TablePrint(String str) {

        //System.out.printf("%10s", str);
        System.out.print(str + "\t");
    }

    /**
     * print by 10s int
     * @param i
     */
    /*
    private static void TablePrint(int i) {

        TablePrint(Integer.toString(i));
    }
    */

}
