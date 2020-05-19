package tools;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class WriteExcel {

    HSSFWorkbook workbook;
    File filexls;



    public WriteExcel(String name){
        workbook =  new HSSFWorkbook();
        filexls = new File(name);
    }
    public HSSFSheet getsheet(String sheetname){
        return workbook.createSheet(sheetname);
    }
    public HSSFRow getrow(HSSFSheet sheet , int rownum){
        return sheet.createRow(rownum);
    }
    public HSSFCell getcell(HSSFRow row,int colnum){
        return row.createCell(colnum);
    }
    public void makeCellValue(HSSFCell cell,String info){
        cell.setCellValue(info);
        try {
            this.saveFile();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public void easymakeCellValue(String sheetname,int rownum,int colnum,String info){
        workbook.createSheet(sheetname).createRow(rownum).createCell(colnum).setCellValue(info);
        try{
            this.saveFile();
        }catch(IOException e){
            e.printStackTrace();
        }
    }

    void saveFile() throws IOException {
        FileOutputStream fileout =  new FileOutputStream(filexls);
        workbook.write(fileout);
        fileout.flush();
        fileout.close();
    }


    public static void main(String[] args) throws ParseException {
        Date now  = new Date();
        SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss -SSS");
        String nowtime = sf.format(now);
        System.out.println(nowtime);
        Date timem = sf.parse(nowtime);
        System.out.println(timem);

//        WriteExcel wt =new WriteExcel("HHE.xls");
//        HSSFSheet sheet = wt.getsheet("sheet1");
//        HSSFRow row = wt.getrow(sheet,2);
//        HSSFCell cell = wt.getcell(row,3);
//        wt.makeCellValue(cell,"123");

//        wt.easymakeCellValue("sheet2",1,1,"nihao");

    }
//    public static void main(String[] args){
//        HSSFWorkbook workbook =  new HSSFWorkbook();
//        HSSFSheet sheet = workbook.createSheet("sheet1");
//        HSSFRow row = sheet.createRow(3);
//        row.createCell(3).setCellValue("a");
//
//        File xlsfile = new File("xxx.xls");
//        FileOutputStream xlsStream = null;
//        try {
//            xlsStream = new FileOutputStream(xlsfile);
//        } catch (FileNotFoundException e) {
//            e.printStackTrace();
//        }
//        try {
//            workbook.write(xlsStream);
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//
//    }
}
