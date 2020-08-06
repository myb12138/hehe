package com.itheima;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class PoiDemo {

    @Test
    public void readExcel() throws IOException {
        //1.获取一个指向文件的workbook
        Workbook workbook = new XSSFWorkbook("D:\\Develop\\黑马所有资料\\项目一(SAAS)_杜宏\\day10-货物导入_出货打印_百万数据报表\\资料\\poi资料\\demo.xlsx");
        //2.获取sheet
        Sheet sheet = workbook.getSheetAt(0);
        //3.循环读取行列数据
        for (int i = 0; i <= sheet.getLastRowNum(); i++) { //读取行
            Row row = sheet.getRow(i);
            for (int j = 0; j < row.getLastCellNum(); j++) { //读取列
                Cell cell = row.getCell(j);
                //读取cell中的数据   数据是有类型的，那么我们在读取时要考虑类型问题
                if (cell != null) {
                    Object obj = getCellValue(cell);
                    System.out.print(obj + "    ");
                }
            }
            System.out.println();
        }
    }

    //判断cell的类型，通过不同的方法获取Cell中的值
    private Object getCellValue(Cell cell) {
        Object obj = null;
        CellType cellType = cell.getCellType(); //cell类型
        Double
        switch (cellType) {
            case STRING:
                obj = cell.getStringCellValue();
                break;
            case BOOLEAN:
                obj = cell.getBooleanCellValue();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) { //判断是否是日期格式
                    obj = cell.getDateCellValue();
                } else {
                    obj = cell.getNumericCellValue();
                }
                break;
        }
        return obj;
    }


    @Test
    public void writeExcel() throws IOException {
        //1.创建workbook
        Workbook workbook = new XSSFWorkbook();
        //2.创建sheet
        Sheet sheet = workbook.createSheet();
        //3.创建row
        Row row = sheet.createRow(1);
        //4.创建cell
        Cell cell = row.createCell(1);
        //5.向cell中写入内容
        cell.setCellValue("传智播客");
        //设置样式
        //- 设置行高与列宽
        row.setHeightInPoints(50);//行高
        sheet.setColumnWidth(1, 20 * 256);//列宽
        //- 设置样式细节
        CellStyle cellStyle = workbook.createCellStyle();
        //  - 居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        //  - 边框
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        //  - 字体
        Font font = workbook.createFont();
        font.setFontName("华文行楷");
        font.setBold(true);
        font.setFontHeightInPoints((short) 26);
        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle); //设置cell样式
        //6.写入到excel
        FileOutputStream fos = new FileOutputStream("d:/aa.xlsx");
        workbook.write(fos);
        fos.flush();
        fos.close();
    }
}
