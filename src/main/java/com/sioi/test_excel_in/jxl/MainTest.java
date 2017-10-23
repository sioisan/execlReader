package com.sioi.test_excel_in.jxl;

import com.sioi.test_excel_in.jxl.ExcelFlow.Type;

public class MainTest {

    public static void main(String[] args) {

        ExcelFlow ef = new ExcelFlow("D:/excel.xls", Type.READ);
        ef.setSheet(0);
        ef.getCells().getPrint();
        ef.close();
    }

}
