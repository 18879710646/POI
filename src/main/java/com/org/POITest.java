package com.org;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @Author HP
 * @Date 2021/12/1 21:23
 * @Version 1.0
 * 其它事与我无关，多看一眼都是愚蠢的。
 * 享有特权而没有力量的人是废物，
 * 受过教育而无影响力的人是一堆一文不值的垃圾。
 */

public class POITest {
    // 获取带excel中的文件（不能获取到一列中既包含文字和数字的数据）
    @Test
    public void Test() throws IOException {
        XSSFWorkbook sheets = new XSSFWorkbook("D:\\Demo2021\\read.xlsx");
        XSSFSheet sheetAt = sheets.getSheetAt(0);
        for (Row cells : sheetAt) {
            for (Cell cell : cells) {
                String value = cell.getStringCellValue();
                System.out.println(value);
            }
        }
        sheets.close();

    }

    // 按要求获取到数据
    @Test
    public void TestLastRow() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook("D:\\Demo2021\\read.xlsx");
        XSSFSheet sheetAt = workbook.getSheetAt(0);
        // 获取第几行中的数据集
            XSSFRow row = sheetAt.getRow(1);
            // 获取数据集中的size
            short lastCellNum = row.getLastCellNum();
            // 进行每一个数据遍历
            for (int j=0;j<lastCellNum;j++){
                // 获取到每一个数据
                String value = row.getCell(j).getStringCellValue();
                System.out.println(value);
            }
    }
    @Test
    // 获取到最后一个数据并且获取到最后一个单元格中的数据
    public void TestLast() throws IOException {
        XSSFWorkbook sheets = new XSSFWorkbook("D:\\Demo2021\\read.xlsx");
        XSSFSheet sheetAt = sheets.getSheetAt(0);
        int lastRowNum = sheetAt.getLastRowNum();
        System.out.println("显示出最后一行的数据是第几行"+lastRowNum);   //2
        XSSFRow row = sheetAt.getRow(lastRowNum-1);
        // 获取到最后一个单元格的号码是多少,也就是一行中总共有多少个单元格
        short lastCellNum = row.getLastCellNum();
        System.out.println("一行中有多少个单元格中拥有数据:"+lastCellNum);
        for (int i=0;i<lastCellNum;i++){
            if (i==(lastCellNum-1)) {
                String value = row.getCell(i).getStringCellValue();
                System.out.println("最后一行中中的最后一个单元格中的数据是："+value);
            }
            else {
                continue;
            }

        }

    }

    // 创建一张表

    @Test
    public void  CreateSheet() throws Exception {
        // 创建一个工作溥
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 创建一个工作表为"demo1"
        XSSFSheet sheet = workbook.createSheet("demo1");
        // 生成第一行
        XSSFRow row = sheet.createRow(0);
        // 为第一行中的第一个单元格添加数据
        row.createCell(0).setCellValue("姓名：");
        row.createCell(1).setCellValue("年龄");
        row.createCell(2).setCellValue("地址");
        // 生成第二行
        XSSFRow row1 = sheet.createRow(1);
        // 为第二行添加数据
        row1.createCell(0).setCellValue("郑建成");
        row1.createCell(1).setCellValue("12");
        row1.createCell(2).setCellValue("江西省赣州市");
        // 生成数据流，并添加ptah
        FileOutputStream outputStream = new FileOutputStream("D:\\itdemo1.xlsx");
        // 数据流写进到工作溥中
        workbook.write(outputStream);
        // 数据更新
        outputStream.flush();
        //关闭
        outputStream.close();
        workbook.close();
    }


    // 获取整张表中的数据
    @Test
    public void  QueryTestdemo() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook("D:\\itdemo1.xlsx");
        XSSFSheet sheetAt = workbook.getSheetAt(0);
        for (Row row : sheetAt) {
            for (Cell cell : row) {
                System.out.println(cell.getStringCellValue());

            }
        }
    }
}
