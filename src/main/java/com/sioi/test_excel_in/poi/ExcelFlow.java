package com.sioi.test_excel_in.poi;

import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * for just read
 * @author sioi
 *
 */
public class ExcelFlow{
    
    private File file;
    private FileInputStream fileInputStream;
    private Workbook workbook = null;
    private int sheetMaxPage;
    private Sheet sheet;
    private String[] title;
    
    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";
    
    
    ExcelFlow(String fileName){
        
        file = new File(fileName);
        try {
            
            String fileType = isExcelFile(file);
            fileInputStream = new FileInputStream((file));
            workbook = fileType.equals(XLS) ? new HSSFWorkbook(fileInputStream) :
                                fileType.equals(XLSX)?new XSSFWorkbook(fileInputStream):null;
            sheetMaxPage = workbook.getNumberOfSheets();
            sheet = workbook.getSheetAt(0);   
            
        } catch (Exception e) {
            e.printStackTrace();
        }
        
        
    }
    
    public void toBean(Object o,String methodName,String value,String method){
        try {
            Class c = Class.forName(o.getClass().getName());
            String setMethod = method + methodName.substring(0,1).toUpperCase() + methodName.substring(1);
            Method [] m = c.getMethods();
            
            for(Method mm : m ){
                if(setMethod.equals(mm.getName())){
                    Class[] params = mm.getParameterTypes();
                    if(params[0].getName().equals("int")){
                        mm.invoke(o,Integer.parseInt(value));
                        break;
                    }else if(params[0].getName().equals("java.lang.String")){
                        mm.invoke(o, value);
                        break;
                    }
                }
            }
            
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    public List<Person> toBean(Person person,int sheetIndex){
        
        List<Person> objectList = new ArrayList<Person>();
        
        this.sheet = workbook.getSheetAt(sheetIndex);
        title = new String [6];        
        int rowCount = 0;
        int columnCount = 0;
        int column = 0;
        boolean flag = false;
        

        for(Row row : sheet){
            column = 0;
            if(rowCount == 0){
                flag = true;
            }else{
                flag = false;
            }
            if(row == null){
                continue;
            }
            Person pp = new Person();
            
            for(Cell cell : row){
                
                if(printByString(cell).equals("")||cell == null){
                    break;
                }else{
                if(flag){
                    title[columnCount++] = printByString(cell);
                }else{
                    toBean(pp, title[column++],printByString(cell),"set");
                }
                }
            }
            rowCount++;
            if(!flag){
                objectList.add(pp);
            }
            
        }
        
        return objectList;
    }
    
    public String printByString(Cell cell){
        
        
        switch(cell.getCellTypeEnum()){
        case STRING : return cell.getStringCellValue();
        case NUMERIC : return  Long.toString(Math.round(cell.getNumericCellValue()));
        default:return cell.toString();
        }
        
    }
    
    public void print(int index){
        this.sheet = workbook.getSheetAt(index);
        
        for(Row row : sheet){
             
            for(Cell cell : row){

               System.out.print(printByString(cell) + "\t");
                
            }
            System.out.println();
        }
        
    }
    
    private String isExcelFile(File file) throws Exception{
        
        if(file.exists()){
            
            if(file.isFile()){
                
                if(file.getName().endsWith("xls")){
                    return XLS;
                    
                }else if(file.getName().equals("xlsx")){
                    return XLSX;
                }
                else{
                    throw new Exception("The file is not excel file,please check it over");
                }
                
            }else{
                throw new Exception("It is not file");
                
            }
            
        }else{
            throw new Exception("File not found.");
        }
        
       }
    
    public static void main(String [] args){
        ExcelFlow ef = new ExcelFlow("D:\\excel.xls");
        List<Person> lp = ef.toBean(new Person(), 0);
        for(Person person : lp)
        System.out.println(person.getId() + " " + person.getAge() + " " + person.getName() + person.getSex() + person.getPhoneNumber());
    }
}