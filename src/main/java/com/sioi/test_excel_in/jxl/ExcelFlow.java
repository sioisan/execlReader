package com.sioi.test_excel_in.jxl;

import java.io.File;
import java.io.FileInputStream;

import org.junit.Test;

import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.Number;

/**
 * @author sioi
 *
 */
public class ExcelFlow {

    /*
     * input
     */
    private Workbook book = null;
    private int maxSheetNumber = 0;
    private Sheet sheet = null;
    private File file = null;

    /*
     * output
     */
    private WritableWorkbook writableWorkbook = null;
    private WritableSheet writableSheet = null;
    private Label label;

    public enum Type {
        READ, NEW, UPDATE
    };

    private int Columns = 0;
    private int Rows = 0;
    private int sheetIndex = 0;

    /**
     * @param fileName
     *            fileName
     * 
     * @param type
     *            what to do
     */
    public ExcelFlow(String fileName, Type type) {

        file = new File(fileName);

        try {
            if (!file.exists() || Type.NEW == type) {
                file.createNewFile();
                writableWorkbook = Workbook.createWorkbook(file);
            } else {
                if (Type.UPDATE == type) {

                    book = Workbook.getWorkbook(new FileInputStream(new File(fileName)));

                    writableWorkbook = Workbook.createWorkbook(new File(fileName), book);

                } else if (Type.READ == type) {
                    book = Workbook.getWorkbook(new File(fileName));
                } else {
                    this.close();
                }

            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        if (book != null) {
            maxSheetNumber = book.getNumberOfSheets();
            sheet = book.getSheet(0);
            Columns = sheet.getColumns();
            Rows = sheet.getRows();
        }
    }

    /************* get sheet *************/

    /**
     * 
     * @param index
     * @return Sheet get Sheet by index
     */
    public Sheet getSheet(int index) {

        if (maxSheetNumber > index && index >= 0) {
            sheet = book.getSheet(index);
            sheetIndex = index;
            return sheet;
        } else {
            System.out.println("工作表索引越界，当前索引为：" + index + "，最大值为：" + (maxSheetNumber - 1) + "。");
            this.close();
            return null;
        }

    }

    /**
     * 
     * @param index
     *            setSheet by index
     */
    public void setSheet(int index) {

        this.getSheet(index);
    }

    /************* get values *************/
    public Cells getCells(Sheet sheet) {

        Cells cells = new Cells(sheet);
        return cells;

    }

    public Cells getCells() {

        return getCells(this.sheet);
    }

    /************* get the number of row and the number of column *************/
    public int getRows() {

        return this.Rows;
    }

    public int getColumns() {

        return this.Columns;
    }

    /************* create *************/

    /**
     * 
     * @param name
     * @param index
     * @return WritableSheet
     */
    public WritableSheet createSheet(String name, int index) {

        writableSheet = writableWorkbook.createSheet(name, index);

        return writableSheet;

    }

    /**
     * 
     * @param indexRow
     * @param indexColumn
     * @param name
     * @return Label
     */
    public Label createLaber(int indexRow, int indexColumn, String name) {

        label = new Label(indexRow, indexColumn, name);
        return label;

    }

    /************* insert *************/

    /**
     * 
     * @param label
     * @param sheet
     */
    public void insertLabel(Label label, WritableSheet sheet) {

        try {
            sheet.addCell(label);
        } catch (Exception e) {

            e.printStackTrace();
        }
    }

    /**
     * 
     * @param sheet
     */
    public void insertLabel(WritableSheet sheet) {

        insertLabel(this.label, sheet);

    }

    public void insertLabel() {

        insertLabel(this.writableSheet);
    }

    /**
     * create and insert label with String
     * 
     * @param indexRow
     * @param indexColumn
     * @param name
     * @param indexSheet
     */
    public void createInsertLabel(int indexRow, int indexColumn, String name, int indexSheet) {

        writableSheet = writableWorkbook.getSheet(indexSheet);
        createLaber(indexRow, indexColumn, name);
        insertLabel();
    }

    public void creaeteInsertLabel(int indexRow, int indexColumn, String name) {

        createInsertLabel(indexRow, indexColumn, name, this.sheetIndex);
    }

    /**
     * insert number
     * 
     * @param indexRow
     * @param indexColumn
     * @param number
     */
    public void insertNum(int indexRow, int indexColumn, double number) {

        Number num = new Number(indexRow, indexColumn, number);
        try {
            writableSheet.addCell(num);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    /************* save file *************/
    /**
     * save
     */
    public void save() {

        try {
            writableWorkbook.write();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /************* close file and release memory *************/
    public void close() {

        try {
            if (writableWorkbook != null) {
                writableWorkbook.close();
            }
            if (book != null) {
                book.close();
            }
            System.gc();

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    /************* test *************/
    @Test
    private static void ReadTest() {

        ExcelFlow ef = new ExcelFlow("D:/excel.xls", Type.READ);
        ef.setSheet(0);
        ef.getCells().getPrint();
        ef.close();

    }

    @Test
    private static void NewTest() {

        ExcelFlow ef = new ExcelFlow("D:\\excel.xls", Type.NEW);

        ef.createSheet("aaa", 0);

        ef.creaeteInsertLabel(0, 9, "aaa");
        ef.save();
        ef.close();

    }

    @Test
    private static void UpdateTest() {

        ExcelFlow ef = new ExcelFlow("D:\\excel.xls", Type.UPDATE);
        ef.setSheet(0);
        ef.creaeteInsertLabel(0, 7, "aab");
        ef.insertNum(2, 3, 99.99);
        ef.save();
        ef.close();
    }

}
