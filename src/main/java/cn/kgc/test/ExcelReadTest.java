package cn.kgc.test;


import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.joda.time.*;
import org.junit.jupiter.api.*;

import javax.lang.model.element.*;
import java.io.*;
import java.util.*;

/**
 * @author songyuhang
 * @create 2020-08-10 16:00
 */
public class ExcelReadTest {

    //路径变量
    String PATH="E:\\workpaceid\\Excel_test\\poi_test\\";

    /*测试03版本读取*/
    @Test
    public void TestRead03() throws Exception {

        //1.获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH+"poi_test信息表03.xls");
        //2.读取工作簿中的内容
        Workbook workbook=new HSSFWorkbook(fileInputStream);
        //3.获取工作中的表
        Sheet sheetAt = workbook.getSheetAt(0); //根据下标获取 也可根据表名查找
        //4.获取表中的行
        Row row1 = sheetAt.getRow(0);
        //5.获取单元格
        Cell cellA1 = row1.getCell(0);
        //6.取出单元格中的值
        // 读取值的时候需要注意表中数据的类型
        System.out.println(cellA1.getStringCellValue());   //getStringCellValue() 获取的是制字符串类型
        //关闭流
        fileInputStream.close();
    }


    /*测试07版本读取*/
    @Test
    public void TestRead07() throws Exception {

        //1.获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH+"poi_test信息表07.xlsx");

        //2.读取工作簿中的内容
        Workbook workbook=new XSSFWorkbook(fileInputStream);

        //3.获取工作簿中的表
        Sheet sheetAt = workbook.getSheetAt(0);

        //4.获取表中的行
        Row row2 = sheetAt.getRow(1);

        //5.获取单元格
        Cell cellB2 = row2.getCell(1);

        //6.取出单元格中的值
        System.out.println(cellB2.getStringCellValue());   //getStringCellValue() 获取的是制字符串类型

        //关闭流
        fileInputStream.close();
    }


    /*读取不同类型的数据*/
    @Test
    public void testCellType() throws Exception {
        //1.获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH+"明细表.xls");
        //2.读取工作簿中的内容
        Workbook workbook=new HSSFWorkbook(fileInputStream);
        //3.获取工作簿中的表
        Sheet sheetAt = workbook.getSheetAt(0);
        //4.获取标题内容
        Row rowTitle = sheetAt.getRow(0);    //第一行 标题
        // 判断标题内容是否为空
        if(rowTitle!=null){
            //得到第一行标题的个数 （获取所有列数）
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum=0;cellNum<cellCount;cellNum++){
                Cell cell = rowTitle.getCell(cellNum);
                if (cell!=null){
                    //获取类型
                    int cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue+" | ");
                }
            }
            System.out.println();
        }
        //5.获取表中的内容
            //获取所有行数
        int rowCount= sheetAt.getPhysicalNumberOfRows();
        for(int rowNum=1;rowNum<rowCount;rowNum++){
            //获取行
            Row rowData = sheetAt.getRow(rowNum);
            //判断获取到每一行是否为空
            if(rowData!=null){
                //读取列
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for(int cellNum=0;cellNum<cellCount;cellNum++){
                    System.out.print("["+(rowNum+1)+","+(cellNum+1)+"]");
                    //获取列 获取单元格
                    Cell cell = rowData.getCell(cellNum);
                    /**
                     * 拿到单元格后并不能知道单元格的类型
                     * 需要去匹配数据类型
                     */
                     if(cell!=null){
                         //获取列类型
                         int cellType = cell.getCellType();
                         String cellVaule="";
                         switch (cellType){
                             case HSSFCell.CELL_TYPE_STRING:   //为字符串类型
                                 System.out.print("[String]");
                                 //获取值
                                  cellVaule = cell.getStringCellValue();
                                  break;
                             case HSSFCell.CELL_TYPE_BOOLEAN:   //为布尔类型
                                 System.out.print("[Boolean]");
                                 //获取值
                                 cellVaule = String.valueOf(cell.getBooleanCellValue());
                                 break;
                             case HSSFCell.CELL_TYPE_BLANK:   //为空
                                 System.out.print("[Blank]");
                                 break;
                             case HSSFCell.CELL_TYPE_NUMERIC:   //为数字 （日期和普通数字）
                                 System.out.print("[NUMERIC]");
                                 //再次判断数字为日期还是普通数字
                                 if(HSSFDateUtil.isCellDateFormatted(cell)){    //日期
                                     System.out.print("[日期]");
                                     Date date = cell.getDateCellValue();
                                     cellVaule = new DateTime(date).toString("yyyy-MM-dd");
                                 }else{                                               //为普通数字
                                     //如果不是日期，防止数字过长(科学计数法)，转换成字符串输出
                                     System.out.print("[数字]");
                                     cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                     cellVaule=cell.toString();
                                 }
                                 break;
                             case HSSFCell.CELL_TYPE_ERROR:   //数据类型错误
                                 System.out.print("[数据类型错误]");
                                 break;
                         }
                         System.out.println(cellVaule);
                     }
                }
            }
        }
        //关流
        fileInputStream.close();
    }
}
