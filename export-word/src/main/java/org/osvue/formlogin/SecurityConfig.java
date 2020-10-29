package org.osvue.formlogin;


import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.List;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.List;


/**
 * @Author: Mr.Han
 * @Description: securityFive
 * @Date: Created in 2020/5/23_10:45
 * @Modified By: THE GIFTED
 */
//@Configuration
//public class SecurityConfig extends WebSecurityConfigurerAdapter {
public class SecurityConfig   {
    public static void main(String[] args) throws Exception {
      word2();
    }

    public static void word2(){
      XWPFDocument doc = new XWPFDocument();
      titleStyle(doc,"项目信息列表");


      // 创建20行7列
      XWPFTable table = doc.createTable(21, 7);
      tableBorderStyle(table);
      // table.set
      List<XWPFTableCell> tableCells1 = table.getRow(0).getTableCells();
      tableTextStyle(tableCells1,0,"编码");
      tableTextStyle(tableCells1,1,"名称");
      tableTextStyle(tableCells1,2,"地址");
      tableTextStyle(tableCells1,3,"电话");
      tableTextStyle(tableCells1,4,"负责人");
      tableTextStyle(tableCells1,5,"类型");
      tableTextStyle(tableCells1,6,"备注");
      for(int i=1;i<21;i++){
        List<XWPFTableCell> tableCells2 = table.getRow(i).getTableCells();
        for(int j=0;j<7;j++){
          tableTextStyle(tableCells2.get(j),i+"->"+j);
        }
      }
      try {
        File f=new File("E:\\tmp\\wordTest\\aaa-"+Math.random()+".docx");
        if(f.exists()==false){
          f.createNewFile();
        }
        FileOutputStream out = new FileOutputStream(f);
        doc.write(out);
        out.close();
      } catch (Exception e) {
        e.printStackTrace();
      }
    }

    private static void tableBorderStyle(XWPFTable table){
      //表格属性
      CTTblPr tablePr = table.getCTTbl().addNewTblPr();
      //表格宽度
      CTTblWidth width = tablePr.addNewTblW();
      width.setW(BigInteger.valueOf(8000));
      //表格颜色
      CTTblBorders borders=table.getCTTbl().getTblPr().addNewTblBorders();
      //表格内部横向表格颜色
      CTBorder hBorder=borders.addNewInsideH();
      hBorder.setVal(STBorder.Enum.forString("single"));
      hBorder.setSz(new BigInteger("1"));
      hBorder.setColor("dddddd");
      //表格内部纵向表格颜色
      CTBorder vBorder=borders.addNewInsideV();
      vBorder.setVal(STBorder.Enum.forString("single"));
      vBorder.setSz(new BigInteger("1"));
      vBorder.setColor("dddddd");
      //表格最左边一条线的样式
      CTBorder lBorder=borders.addNewLeft();
      lBorder.setVal(STBorder.Enum.forString("single"));
      lBorder.setSz(new BigInteger("1"));
      lBorder.setColor("dddddd");
      //表格最左边一条线的样式
      CTBorder rBorder=borders.addNewRight();
      rBorder.setVal(STBorder.Enum.forString("single"));
      rBorder.setSz(new BigInteger("1"));
      rBorder.setColor("dddddd");
      //表格最上边一条线（顶部）的样式
      CTBorder tBorder=borders.addNewTop();
      tBorder.setVal(STBorder.Enum.forString("single"));
      tBorder.setSz(new BigInteger("1"));
      tBorder.setColor("dddddd");
      //表格最下边一条线（底部）的样式
      CTBorder bBorder=borders.addNewBottom();
      bBorder.setVal(STBorder.Enum.forString("single"));
      bBorder.setSz(new BigInteger("1"));
      bBorder.setColor("dddddd");
    }

    private static void tableTextStyle(List<XWPFTableCell> tableCells1,int index,String text){
      tableTextStyle(tableCells1.get(index),text);
    }

    private static void tableTextStyle(XWPFTableCell tableCell,String text){
      XWPFParagraph p0 = tableCell.addParagraph();
      tableCell.setParagraph(p0);
      XWPFRun r0 = p0.createRun();
      // 设置字体是否加粗
//        r0.setBold(true);
      r0.setFontSize(12);
      // 设置使用何种字体
      r0.setFontFamily("Helvetica Neue");
      // 设置上下两行之间的间距
      r0.setTextPosition(12);
      r0.setColor("333333");
      r0.setText(text);
    }

    private static void titleStyle(XWPFDocument doc,String title){
      XWPFParagraph p1 = doc.createParagraph();
      // 设置字体对齐方式
      p1.setAlignment(ParagraphAlignment.CENTER);
      p1.setVerticalAlignment(TextAlignment.TOP);
      // 第一页要使用p1所定义的属性
      XWPFRun r1 = p1.createRun();
      // 设置字体是否加粗
      r1.setBold(true);
      r1.setFontSize(20);
      // 设置使用何种字体
      r1.setFontFamily("Courier");
      // 设置上下两行之间的间距
      r1.setTextPosition(20);
      r1.setText(title);
    }
  }


