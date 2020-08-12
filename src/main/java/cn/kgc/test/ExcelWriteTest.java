package cn.kgc.test;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.*;
import org.apache.poi.xssf.usermodel.*;
import org.joda.time.*;
import org.junit.jupiter.api.*;
import java.io.*;

/**
 * @author songyuhang
 * @create 2020-08-10 16:00
 */
public class ExcelWriteTest {

    //构建路径
    String PATH="E:\\workpaceid\\Excel_test\\poi_test";

    /*测试Excel03版本*/
    @Test
    public void testWrite03() throws Exception {
        //1.创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet=workbook.createSheet("信息表");
        //3.创建一个行
        Row row1=sheet.createRow(0);
        //4.创建一个单元格 (构成了 A,1 这个单元格)
        Cell cellA1 = row1.createCell(0);
        //5.向单元格中填写数据
         cellA1.setCellValue("姓名");
            //(构成 A,2 坐标)
        Cell cellB1 = row1.createCell(1);
        cellB1.setCellValue("时间");
            //创建第二行 并传入数据
         Row row2 =sheet.createRow(1);
         Cell cellA2 = row2.createCell(0);
         cellA2.setCellValue("张三");
         Cell cellB2 = row2.createCell(1);
         String time= new DateTime().toString("yyyy-MM-DD");
         cellB2.setCellValue(time);

        //6.生成一张表（IO 流） 03版本后缀名使用xls
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "信息表03.xls");
        //通过流输出输出
        workbook.write(fileOutputStream);
        //关闭流
        fileOutputStream.close();
        System.out.println("信息表03生成完毕！");
    }


    /*测试Excel07版本*/
    @Test
    public void testWrite07() throws Exception {
        //1.创建一个工作簿 使用excel操作的 workBook都能操作
        Workbook workbook = new XSSFWorkbook();
        //2.创建一个工作表 表中的设置
        Sheet sheet=workbook.createSheet("信息表");
        //3.创建一个行
        Row row1=sheet.createRow(0);
        //4.创建一个单元格 (构成了 A,1 这个单元格)
        Cell cellA1 = row1.createCell(0);
        //5.向单元格中填写数据
        cellA1.setCellValue("姓名");
        //(构成 A,2 坐标)
        Cell cellB1 = row1.createCell(1);
        cellB1.setCellValue("时间");
        //创建第二行 并传入数据
        Row row2 =sheet.createRow(1);
        Cell cellA2 = row2.createCell(0);
        cellA2.setCellValue("李四");
        Cell cellB2 = row2.createCell(1);
        String time= new DateTime().toString("yyyy-MM-DD");
        cellB2.setCellValue(time);

        //6.生成一张表（IO 流） 03版本后缀名使用xls
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "信息表07.xlsx");
        //通过流输出输出
        workbook.write(fileOutputStream);
        //关闭流
        fileOutputStream.close();
        System.out.println("信息表07生成完毕！");
    }


    /**
     *测试03版本大数据写入
     * 优点：过程中写入缓存，不操作磁盘，最后一次写入磁盘，速度快
     * 缺点：最多写入65536行，否则抛出异常
     */
    @Test
    public void testWrite03BigData() throws Exception {
        //计算时间 开始时间
        long begin=System.currentTimeMillis();
        //创建工作薄
        Workbook workbook=new HSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int RowNum=0;RowNum<65536;RowNum++){        //RowNum 行数
            Row row = sheet.createRow(RowNum);          //每一行
            for ( int CellNum=0;CellNum<10;CellNum++){
                Cell cell = row.createCell(CellNum);     //每一个单元格
                cell.setCellValue(CellNum+1);
            }
        }
        System.out.println("over!");
        //生成一张表
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "bigData03.xls");
        //输出
        workbook.write(fileOutputStream);
        //关流
        fileOutputStream.close();
        //结束时间
        long end=System.currentTimeMillis();
        System.out.println((double)(end-begin)/1000);
    }


    /**
     *测试07版本大数据写入
     * 优点：可以写入较大的数据量
     * 缺点：写数据的速度慢，消耗内存，如超过100万条数据，也会发生内存溢出现象
     */
    @Test
    public void testWrite07BigData() throws Exception {
        //计算时间 开始时间
        long begin=System.currentTimeMillis();
        //创建工作薄
        Workbook workbook=new XSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int RowNum=0;RowNum<65537;RowNum++){        //RowNum 行数
            Row row = sheet.createRow(RowNum);          //每一行
            for ( int CellNum=0;CellNum<10;CellNum++){
                Cell cell = row.createCell(CellNum);     //每一个单元格
                cell.setCellValue(CellNum+1);
            }
        }
        System.out.println("over!");
        //生成一张表
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "bigData07.xlsx");
        //输出
        workbook.write(fileOutputStream);
        //关流
        fileOutputStream.close();
        //结束时间
        long end=System.currentTimeMillis();
        System.out.println((double)(end-begin)/1000);

    }


    /**
     *测试大文件写SXSSF
     *优点：可以写入非常大的数据量，如100万条甚至更多，写数据数据快，占用更少内存
     * 注意：过程中会产生临时文件，需要清理临时文件
     *       默认有100条数据被存入内存中，如果超过这个数量，则最前面的数据被写入临时文件
     *       自定义内存中的数量，可以使用new SXXFWorkBook（数量）
     */
    @Test
    public void testWrite07BigDatas() throws Exception {
        //计算时间 开始时间
        long begin=System.currentTimeMillis();
        //创建工作薄
        Workbook workbook=new SXSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int RowNum=0;RowNum<65535;RowNum++){        //RowNum 行数
            Row row = sheet.createRow(RowNum);          //每一行
            for ( int CellNum=0;CellNum<10;CellNum++){
                Cell cell = row.createCell(CellNum);     //每一个单元格
                cell.setCellValue(CellNum+1);
            }
        }
        System.out.println("over!");
        //生成一张表
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "bigData07S.xlsx");
        //输出
        workbook.write(fileOutputStream);
        //关流
        fileOutputStream.close();
        //清除临时文件
        ((SXSSFWorkbook)workbook).dispose();
        //结束时间
        long end=System.currentTimeMillis();
        System.out.println((double)(end-begin)/1000);
    }

}
