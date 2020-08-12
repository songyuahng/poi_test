package cn.kgc.test;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.*;

import javax.lang.model.element.*;
import java.io.*;

/**
 * @author songyuhang
 * @create 2020-08-10 19:22
 */
public class TestFormula {
   //路径变量
    String PATH="E:\\workpaceid\\Excel_test\\poi_test\\";
    /*测试计算公式*/
    @Test
    public void testFormula() throws Exception {

        //1.创建流
        FileInputStream fileInputStream = new FileInputStream(PATH+"计算公式.xls");
        //2.获取工作簿
        Workbook Workbook = new HSSFWorkbook(fileInputStream);
        //3.获取表
        Sheet sheetAt = Workbook.getSheetAt(0);
        //4.获取计算公式的单元格
        Row row = sheetAt.getRow(5);
        Cell cell = row.getCell(0);
        //5.拿到计算公式eval
        FormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook)Workbook);
        //6.输出单元格的公式
            //获取类型
        int cellType = cell.getCellType();
            //判断
        switch (cellType){
            case Cell.CELL_TYPE_FORMULA:   //公式
                String cellFormula = cell.getCellFormula();
                //进行计算
                CellValue evaluate = formulaEvaluator.evaluate(cell);
                //得到计算结果的值
                String value = evaluate.formatAsString();
                System.out.println(value);
                break;
        }
    }
}
