package tools;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;

/**
 * 封装读excel表
 */
public class ReadExcel {
    Workbook workbook;

    /**
     * 构造函数传递需要读取的文件名称
     * @param path
     * @throws IOException
     * @throws InvalidFormatException
     */
    public ReadExcel(String path) throws IOException, InvalidFormatException {
        File file = new File(path);
        workbook = WorkbookFactory.create(file);
    }

    /**
     * 获得文件一共多少sheet页
     * @return
     */
    public int getSheetNum(){
        return workbook.getNumberOfSheets();
    }

    /**
     * 获得文件sheet页的名称（传递需要获取名称的sheet序列）
     * @param num
     * @return
     */
    public String getSheetname(int num){
        return workbook.getSheetName(num);
    }

    /**
     * 通过名称和序列获得sheet
     * @param num
     * @return
     */
    public Sheet getSheet(int num){
        return workbook.getSheetAt(num);
    }
    public Sheet getSheet( String  name) {
        return workbook.getSheet(name);
    }

    /**
     * 获得行数（将需要查询的sheet传递进去）
     * @param sheet
     * @return
     */
    public int getRowNum(Sheet sheet){
        return sheet.getLastRowNum()+1;
    }

    /**
     * 获得行，传递sheet和要等到行的行数
     * @param sheet
     * @param num
     * @return
     */
    public Row getRow(Sheet sheet, int num){
        return sheet.getRow(num);
    }

    /**
     * 获得列数，通过需要获得列数的行
     * @param row
     * @return
     */
    public int getColNum(Row row){
        return row.getLastCellNum();
    }

    /**
     * 对行和列数获得值
     * @param row
     * @param col
     * @return
     */
    public String getCellValue(Row row,int col){
        //判断表格类型是String类型，可以直接取String类型
        if(CellType.STRING == row.getCell(col).getCellTypeEnum()){
            return row.getCell(col).getStringCellValue();

        }else if(CellType.NUMERIC == row.getCell(col).getCellTypeEnum()){
            //如果是数字类型，再判断是日期类型还是数字类型，poi中将日期也看做是数字类型
            if(HSSFDateUtil.isCellDateFormatted(row.getCell(col))){
                //如果日期类型就转换为字符串
                SimpleDateFormat dateformat =  new SimpleDateFormat("yyyy-MM-dd");
                System.out.println(row.getCell(col).getCellTypeEnum());
                return dateformat.format(row.getCell(col).getDateCellValue());
            }else{
                row.getCell(col).setCellType(CellType.STRING);
                return row.getCell(col).getStringCellValue();//将得到的数字格式转换成String
            }

        }else{
            return "Null";
        }

//        Cell cell = row.getCell(col);
//        if(HSSFDateUtil.isCellDateFormatted(cell)){
//            SimpleDateFormat dateformat =  new SimpleDateFormat("yyyy/MM/dd");
//            return dateformat.format(cell);
//        }else{
//            cell.setCellType(CellType.STRING);
//            return cell.getStringCellValue();
//        }


    }
    /**
     * 获得值，传递sheetname 和行数和列数即可获得相应值，(按照从0开始数)
     * @param sheetname
     * @param row
     * @param col
     * @return
     */
    public String getEasyValue(String sheetname,int row,int col){
        Sheet sheet = workbook.getSheet(sheetname);
        return sheet.getRow(row).getCell(col).getStringCellValue();
    }
    /**
     * 为案例提供Object[][]类型数据，为@DataProvider使用
     * @param sheetname 读取的sheet名
     * @param startRow 开始行（指定哪一行就从那一行开始读）（按照从1开始读）
     * @param endRow 结束行
     * @param startCol 开始列
     * @param endCol 结束列
     * @return
     */
    public Object[][] getBatchValues(String sheetname, int startRow, int endRow, int startCol, int endCol)  {
        //调用自己的getSheet方法
        Sheet sheet = this.getSheet(sheetname);
        Object[][] datas = new Object[endRow - startRow + 1][endCol - startCol + 1];//确定几行几列
        for (int i = startRow; i <= endRow; i++) {
            Row row = sheet.getRow(i - 1);
            for (int j = startCol; j <= endCol; j++) {
                Cell ge=row.getCell(j - 1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);//先找到这个格子
                ge.setCellType(CellType.STRING);//将格子类型转换为String，否则，excel中是数字会报错
                String value = ge.getStringCellValue();
                datas[i - startRow][j - startCol] = value;
            }
        }
        return datas;
    }
}

//        File excelfile = new File("淘宝手机号归属查询案例.xlsx");
//
//        Workbook workbook = WorkbookFactory.create(excelfile);
//
//        int sheetnum = workbook.getNumberOfSheets();
//        System.out.println("总计"+sheetnum+"个sheett页");
//
//        String sheetname =  workbook.getSheetName(0);
//        System.out.println(sheetname);
//
//        Sheet sheet = workbook.getSheet("sheet1") ;
//        int rownum =  sheet.getLastRowNum()+1;
//        System.out.println("行数"+rownum);
//
//        Row row = sheet.getRow(0);
//        System.out.println(row.getLastCellNum());
//
////        double value =  row.getCell(1).getNumericCellValue();
////        java.text.DecimalFormat formatter = new java.text.DecimalFormat("########");
////        String str = formatter.format(row.getCell(1).getNumericCellValue());
////        System.out.println(str);
//        DecimalFormat formatter =  new DecimalFormat("########");
//        for(int i = 0;i<row.getLastCellNum();i++){
//            if(CellType.NUMERIC==row.getCell(i).getCellTypeEnum()){
//                System.out.println(formatter.format(row.getCell(i).getNumericCellValue()));
//            }else if(CellType.STRING == row.getCell(i).getCellTypeEnum()){
//                System.out.println(row.getCell(i).getStringCellValue());
//            }
//        }
//        CellType type  =row.getCell(1).getCellTypeEnum();
//        System.out.println(type);


