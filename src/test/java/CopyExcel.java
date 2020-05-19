import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import tools.ReadExcel;

import java.io.IOException;
import java.text.SimpleDateFormat;

public class CopyExcel {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        ReadExcel excel =  new ReadExcel("案例一阶段（导入）.xlsx");
        //获得sheet的数量8
        int sheetnum = excel.getSheetNum();
        sheetnum = 2;

//        Sheet sheet  = excel.getSheet(0);
//        Row row = excel.getRow(sheet,1);
//
//        Cell cell = row.getCell(12);
//        System.out.println(cell.getDateCellValue());
//        SimpleDateFormat dateformat = new SimpleDateFormat("yyyy-MM-dd");
//        System.out.println(dateformat.format(cell.getDateCellValue()));




        for(int i = 0;i<sheetnum;i++){
            Sheet sheet = excel.getSheet(i);
            int rowno = excel.getRowNum(sheet);//获得每个sheet页的行数

            for(int j=0;j<rowno;j++){
                Row row = excel.getRow(sheet,j);
                int colnum = excel.getColNum(row);//获得每行有多少列
//                System.out.println(colnum);
                for(int k = 0;k<colnum;k++){
                     System.out.println(excel.getCellValue(row,k));
                }
            }
        }
    }
}
