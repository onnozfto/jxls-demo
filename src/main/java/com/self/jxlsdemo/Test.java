package com.self.jxlsdemo;

import com.self.jxlsdemo.config.JxlsConfig;
import com.self.jxlsdemo.pojo.Department;
import com.self.jxlsdemo.pojo.Person;
import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jxls.common.Context;
import org.jxls.reader.ReaderBuilder;
import org.jxls.reader.ReaderConfig;
import org.jxls.reader.XLSReadStatus;
import org.jxls.reader.XLSReader;
import org.jxls.transform.poi.PoiUtil;

/**
 * @Description
 * @Author will
 * @Date 2019/08/22
 */
public class Test {


  public static void main(String[] args) throws Exception {
    eachDemo();
  }

  /**
   * jx:each(items="person" var="p" lastCell="C2")
   */
  private static void eachDemo() throws Exception {

    List<Person> list = new ArrayList<>();
    list.add(getPerson("福州", "张三", 8));
    list.add(getPerson("龙岩", "刘武", 9));
    list.add(getPerson("宁德", "张三", 10));
    list.add(getPerson("莆田", "李四", 11));
    list.add(getPerson("福州", "王五", 12));
    JxlsBuilder.getBuilder("each.xlsx")
        .out("target/excel/each.xlsx")
        .putVar("person", list)
        .putVar("scale", 0.5)
        .putVar("image", "aa.jpg")
        .build();

  }

  /**
   * jx:grid(lastCell="A2" headers="headers" data="data" areas=[A1:A1,A2:A2] formatCells="String:A2") areas:grid指令应用的模板区域， formatCells:单元格格式化
   * excel模板配置${header} 标题 ${cell} 行的内容
   */
  private static void gridDemo() throws Exception {
    List<Person> list = new ArrayList<>();
    list.add(getPerson("福州", "张三", 8));
    list.add(getPerson("龙岩", "刘武", 9));
    list.add(getPerson("宁德", "张三", 10));
    list.add(getPerson("莆田", "李四", 11));
    list.add(getPerson("福州", "王五", 12));
    //固定写法必须put进去headers,data
    JxlsBuilder.getBuilder("grid.xlsx")
        .out("target/excel/grid.xlsx")
        .configDynamicGrid("name,age,address")
        .putVar("headers", Arrays.asList("姓名", "年龄", "地址"))
        .putVar("data", list)
        .build();
  }

  private static Person getPerson(String address, String name, Integer age) {
    Person person = new Person();
    person.setAddress(address);
    person.setName(name);
    person.setAge(age);
    return person;
  }

  /**
   * mergeCells
   */
  private static void mergeDemo() throws Exception {
    JxlsBuilder.getBuilder("merge.xlsx")
        .out("target/excel/meger.xlsx")
        .putVar("row", 1)
        .build();
  }

  /**
   * 代码配置jxls注释方式实现excel导出
   */
  private static void javaCommentDemo() throws Exception {
    InputStream in = new FileInputStream(new File(JxlsConfig.getTemplateRoot() +
        File.separator + "test.xlsx"));
    OutputStream out = new FileOutputStream(new File("target/excel/test.xlsx"));
    Workbook wb = new XSSFWorkbook(in);
    Sheet sheet = wb.getSheetAt(0);
    Row row = sheet.getRow(0);
    Cell cell = row.getCell(0);
    Cell cell1 = row.getCell(1);
    cell1.setCellValue("${name}");
    PoiUtil.setCellComment(cell, "jx:area(lastCell=\"D20\")", "will", null);
    ByteArrayOutputStream temp = new ByteArrayOutputStream();
    wb.write(temp);

    Context ctx = new Context();
    ctx.putVar("name", "测试jxls无模板生成！");
    JxlsBuilder.getBuilder(new ByteArrayInputStream(temp.toByteArray()))
        .out(out).putVar("name", "测试jxls无模板生成！").build();
  }

  /**
   * 无模板生成excel
   * @throws Exception
   */
  private static void notemplatedemo() throws Exception {
    XSSFWorkbook wb = new XSSFWorkbook();
    XSSFSheet sheet = wb.createSheet("sheet");
    XSSFRow row = sheet.createRow(0);
    XSSFCell cell = row.createCell(0, CellType.STRING);
    XSSFCellStyle cs = wb.createCellStyle();

    XSSFFont font = wb.createFont();//必须通过wb创建字体
    font.setFontName("黑体");
    font.setBold(true);
    font.setItalic(true);
    font.setColor(IndexedColors.BLUE.index);
    cs.setFont(font);
    cs.setFillPattern(FillPatternType.SOLID_FOREGROUND );
    cs.setFillForegroundColor(IndexedColors.RED.getIndex());
    cell.setCellValue("${name}");
    cell.setCellStyle(cs);

    PoiUtil.setCellComment(cell, "jx:area(lastCell=\"B2\")", "will", null);

    ByteArrayOutputStream baos = new ByteArrayOutputStream();
    wb.write(baos);

    InputStream in = new ByteArrayInputStream(baos.toByteArray());
    JxlsBuilder.getBuilder(in).out("target/excel/noTemplate.xlsx")
        .putVar("name", "测试无模板方式生成excel")
        .build();
  }

  /**
   * 导入excel数据映射到javabean对象
   * @throws Exception
   */
  private  static void importExcel()throws Exception {
    Test test = new Test();
    InputStream inputXML = new BufferedInputStream(test.getClass().getResourceAsStream("/xml_config/config.xml"));
    XLSReader mainReader = ReaderBuilder.buildFromXML( inputXML );
    InputStream inputXLS = new BufferedInputStream(test.getClass().getResourceAsStream("/xml_config/import.xlsx"));
    Department department = new Department();
    List persons = new ArrayList();
    Map beans = new HashMap();
    beans.put("department", department);
   // beans.put("list", persons);
    ReaderConfig.getInstance().setSkipErrors( true );//skip 错误
    ReaderConfig.getInstance().setUseDefaultValuesForPrimitiveTypes( true );//出现错误基本类型返回默认值
    XLSReadStatus readStatus = mainReader.read( inputXLS, beans);
    System.out.println(readStatus.getReadMessages());
  }


}
