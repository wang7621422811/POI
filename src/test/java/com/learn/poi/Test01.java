package com.learn.poi;

import com.learn.poi.handler.SheetHandler;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.*;
import java.util.Iterator;

import static javax.xml.bind.JAXBIntrospector.getValue;

/**
 * @author: webin
 * @date: 2020/3/8 20:32
 * @description:
 * @version: 0.0.1
 */
public class Test01 {


    @Test
    public void test01() throws Exception {
        //1.创建workbook工作簿
        Workbook wb = new XSSFWorkbook();
        //2.创建表单Sheet
        Sheet sheet = wb.createSheet("test");
        //3.文件流
        FileOutputStream fos = new FileOutputStream("C:\\Users\\weibin\\learWorkspace\\POI-Learn\\file\\test.xlsx");
        //4.写入文件
        wb.write(fos);
        fos.close();
    }

    //测试创建单元格
    @Test
    public void test02() throws Exception {
        //1.创建workbook工作簿
        Workbook wb = new XSSFWorkbook();
        //2.创建表单Sheet
        Sheet sheet = wb.createSheet("test");
        //3.创建行对象，从0开始
        Row row = sheet.createRow(3);
        //4.创建单元格，从0开始
        Cell cell = row.createCell(0);
        //5.单元格写入数据
        cell.setCellValue("测试");
        //6.文件流
        FileOutputStream fos = new FileOutputStream("C:\\Users\\weibin\\learWorkspace\\POI-Learn\\file\\test.xlsx");
        //7.写入文件
        wb.write(fos);
        fos.close();
    }

    @Test
    public void test03() throws Exception {
        //1.创建workbook工作簿
        Workbook wb = new XSSFWorkbook();
        //2.创建表单Sheet
        Sheet sheet = wb.createSheet("test");
        // 创建单元格格式对象
        CellStyle style = wb.createCellStyle();
        //设置边框
        style.setBorderBottom(BorderStyle.DASH_DOT);//下边框
        style.setBorderTop(BorderStyle.HAIR); //上边框
        //设置字体
        Font font = wb.createFont();//创建字体对象
        font.setFontName("华文行楷");//设置字体
        font.setFontHeightInPoints((short) 28);//设置字号
        style.setFont(font);

        //3.创建行对象，从0开始
        Row row = sheet.createRow(3);
        //设置宽高
        sheet.setColumnWidth(0, 31 * 256);//设置第一列的宽度是31个字符宽度
        row.setHeightInPoints(50);//设置行的高度是50个点
        //设置居中显示
        style.setAlignment(HorizontalAlignment.CENTER);//水平居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中

        //4.创建单元格，从0开始
        Cell cell = row.createCell(0);
        //5.单元格写入数据
        cell.setCellValue("测试");
        //设置单元格样式
        cell.setCellStyle(style);
        //合并单元格
        CellRangeAddress region = new CellRangeAddress(0, 3, 0, 2);
        sheet.addMergedRegion(region);
        //6.文件流
        FileOutputStream fos = new FileOutputStream("C:\\Users\\weibin\\learWorkspace\\POI-Learn\\file\\test03.xlsx");
        //7.写入文件
        wb.write(fos);
        fos.close();
    }

    /**
     * 绘制图形
     */
    @Test
    public void test04() throws IOException {
        //1. 创建workbook工作簿
        Workbook wb = new XSSFWorkbook();
        //2. 创建sheet表单
        Sheet sheet = wb.createSheet("TEST");
        //3. 读取图片流
        FileInputStream fis = new FileInputStream("E:\\JAVA相关\\基于SaaS平台的iHRM实战开发\\08-员工管理及POI\\01-员工管理及POI入门\\资源\\资源\\Excel相关\\logo.jpg");
        byte[] bytes = IOUtils.toByteArray(fis);
        fis.read(bytes);
        // 向Excel添加一张图片并返回图片的下标
        int pictureIndex = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
        CreationHelper creationHelper = wb.getCreationHelper();
        //5. 绘制一个图片
        Drawing<?> drawingPatriarch = sheet.createDrawingPatriarch();
        //6.创建锚点绘制图片坐标
        ClientAnchor clientAnchor = creationHelper.createClientAnchor();
        clientAnchor.setCol1(0);
        clientAnchor.setRow1(0);
        // 创建图片
        Picture picture = drawingPatriarch.createPicture(clientAnchor, pictureIndex);
        picture.resize();
        //8. 输出文件流
        FileOutputStream fos = new FileOutputStream("C:\\Users\\weibin\\learWorkspace\\POI-Learn\\file\\test04.xlsx");
        wb.write(fos);
        fos.close();
    }

    @Test
    public void test05() throws IOException {
        Workbook wb = new XSSFWorkbook("E:\\JAVA相关\\基于SaaS平台的iHRM实战开发\\08-员工管理及POI\\02-POI报表的高级应用\\资源\\百万数据报表\\demo.xlsx");
        //2. 获取sheet从0开始
        Sheet sheet = wb.getSheetAt(0);
        //获取总行数
        int lastRowNum = sheet.getLastRowNum();
        Row row = null;
        Cell cell = null;

        //循环所有行
        for (int rowNum = 0; rowNum <=sheet.getLastRowNum(); rowNum++) {
            row = sheet.getRow(rowNum);
            StringBuilder sb = new StringBuilder();
            //循环每行中的所有单元格
            for(int cellNum = 0; cellNum <row.getLastCellNum();cellNum++) {
                cell = row.getCell(cellNum);
                sb.append(getValue(cell)).append("-");
            }
            sb.append(getValue(cell)).append("\n");
            System.out.println(sb.toString());
        }


    }

    //获取数据
    private static Object getValue(Cell cell) {
        Object value = null;
        switch (cell.getCellType()) {
        case STRING: //字符串类型
        value = cell.getStringCellValue();
        break;
        case BOOLEAN: //boolean类型
        value = cell.getBooleanCellValue();
        break;
        case NUMERIC: //数字类型（包含日期和普通数字）
        if(DateUtil.isCellDateFormatted(cell)) {
        value = cell.getDateCellValue();
        }else{
        value = cell.getNumericCellValue();
        }
        break;
        case FORMULA: //公式类型
        value = cell.getCellFormula();
        break;
        default:
        break;
        }
        return value;
    }

    /**
     * 使用事件模型读取数据
     */
    @Test
    public void test5() throws OpenXML4JException, IOException, SAXException {
        String path = "E:\\JAVA相关\\基于SaaS平台的iHRM实战开发\\08-员工管理及POI\\02-POI报表的高级应用\\资源\\百万数据报表\\demo.xlsx";
        // 根据excel报表获取OPCpackage
        OPCPackage opc = OPCPackage.open(path, PackageAccess.READ);
        // 创建XSSFReader
        XSSFReader reader = new XSSFReader(opc);
        //获取SharedStringTable对象
        SharedStringsTable stringsTable = reader.getSharedStringsTable();
        //获取StyleTable对象
        StylesTable stylesTable = reader.getStylesTable();
        //创建Sax的xmlReader对象
        XMLReader xmlReader = XMLReaderFactory.createXMLReader();
        //注册事件处理器
        XSSFSheetXMLHandler xssfSheetXMLHandler = new XSSFSheetXMLHandler(stylesTable, stringsTable, new SheetHandler(), false);
        xmlReader.setContentHandler(xssfSheetXMLHandler);
        //逐行读取
        XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator) reader.getSheetsData();
        while (sheetIterator.hasNext()) {
            InputStream stream = sheetIterator.next();
            InputSource source = new InputSource(stream);
            xmlReader.parse(source);
        }
    }


}
