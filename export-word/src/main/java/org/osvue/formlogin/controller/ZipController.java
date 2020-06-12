package org.osvue.formlogin.controller;

import com.lowagie.text.Cell;
import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Element;
import com.lowagie.text.Font;
import com.lowagie.text.PageSize;
import com.lowagie.text.Paragraph;
import com.lowagie.text.Phrase;
import com.lowagie.text.Table;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.rtf.RtfWriter2;
import com.lowagie.text.rtf.style.RtfFont;
import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.tools.zip.ZipOutputStream;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import org.osvue.formlogin.domain.FileBean;
import org.osvue.formlogin.utils.DocStyleUtils;
import org.osvue.formlogin.utils.ZipUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

/**
 * 使用itext 导出word @Author: Mr.Han @Description: securityFive @Date: Created in
 * 2020/6/11_17:08 @Modified By: THE GIFTED
 */
@Controller
public class ZipController {
  // 如果作为压缩包和另外两个附件一起导出则直接先生成word文档到
  // D:\workspace101\mian\WebContent\WEB-INF\temDictionary文件夹下。

  public void createDocContext(String file, String contextString)
      throws DocumentException, IOException {
    // 设置纸张大小
    Document document = new Document(PageSize.A4);
    // 建立一个书写器，与document对象关联
    RtfWriter2.getInstance(document, new FileOutputStream(file));
    document.open();

    document.setMargins(90f, 90f, 72f, 72f);
    // 设置标题字体样式：方正小标宋_GBK、二号、粗体
    // Font tFont= DocStyleUtils.setFontStyle("方正小标宋_GBK", 22, Font.BOLD);
    RtfFont tFont = new RtfFont("方正小标宋_GBK", 22, Font.BOLD, Color.BLACK);
    // 设置正文字体样式:仿宋_GB2312、三号、常规
    // Font cFont = DocStyleUtils.setFontStyle("仿 宋_GB2312", 16f, Font.NORMAL);
    // 设置中文字体
    BaseFont bfChinese = BaseFont.createFont();
    //        BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
    // 标题字体风格
    Font titleFont = new Font(bfChinese, 12, Font.BOLD);
    // 正文字体风格
    Font contextFont = new Font(bfChinese, 10, Font.NORMAL);
    // -----------
    // 构建标题
    Font cFont = DocStyleUtils.setFontStyle("仿 宋", 16f, Font.NORMAL);
    Paragraph title =
        DocStyleUtils.setParagraphStyle(
            "配电线路故障周报第十二期（03月27日00时—04月03日00时）", tFont, 12f, Paragraph.ALIGN_CENTER);
    // 构建[接收单位]
    Paragraph jsdw =
        DocStyleUtils.setParagraphStyle(
            "【接收单位】", /*DocStyleUtils.setFontStyle("黑 体", 18f, Font.BOLD)*/
            new RtfFont("黑 体", 18f, Font.NORMAL, Color.BLACK),
            32f,
            1.2f);
    // [接收单位内容]
    StringBuffer sb = new StringBuffer();
    sb.append("公司各单位,国网辽宁电科院");
    Paragraph jsdwnr = DocStyleUtils.setParagraphStyle(sb.toString(), cFont, 32f, 1.2f);
    // 构建【通报内容】
    Paragraph tbnr =
        DocStyleUtils.setParagraphStyle(
            "【通报内容】", /* DocStyleUtils.setFontStyle("黑 体", 18f, Font.BOLD)*/
            new RtfFont("黑 体", 18f, Font.NORMAL, Color.black),
            32f,
            1.2f);
    tbnr.setLeading(30f);
    // 一、总体情况
    Paragraph ztqk =
        DocStyleUtils.setParagraphStyle(
            " 一、总体情况", DocStyleUtils.setFontStyle("方正黑体_GBK", 16f, Font.BOLD), 32f, 1.2f);
    StringBuffer finalResult = new StringBuffer();
    finalResult.append(
        "2月28日0时至24时，北京市新增1例新冠肺炎确诊病例，为确诊病例的密切接触者，已送至定点医疗机构进行救治。报告新增疑似病例9例、死亡病例1例、密切接触者16人。治愈出院患者14例，分别从市区两级定点医院出院。其中有8名男性，6名女性，年龄最小的36岁，最大的70岁。\n"
            + "\n"
            + "截至2月28日24时,累计确诊病例411例，治愈出院病例271例，死亡病例8例。现有疑似病例33例。累计确定密切接触者2683人，其中508人尚在隔离医学观察中。确诊病例中东城区13例、西城区53例、朝阳区71例、海淀区62例、丰台区42例、石景山区14例、门头沟区3例、房山区16例、通州区19例、顺义区10例、昌平区29例、大兴区39例、怀柔区7例、密云区7例、延庆区1例，外地来京人员25例。平谷区尚未有病例。");
    Paragraph p1Content = DocStyleUtils.setParagraphStyle(finalResult.toString(), cFont, 32f, 1.2f);
    // --------------------------------------
    // Paragraph title = new Paragraph("标题");
    // 设置标题格式对齐方式
    title.setAlignment(Element.ALIGN_CENTER);
    title.setFont(titleFont);

    Paragraph context = new Paragraph(contextString);
    context.setAlignment(Element.ALIGN_LEFT);
    context.setFont(contextFont);
    // 段间距
    context.setSpacingBefore(3);
    // 设置第一行空的列数
    context.setFirstLineIndent(20);

    // 设置Table表格,创建一个三列的表格
    Table table = new Table(3);
    int width[] = {25, 25, 50}; // 设置每列宽度比例
    table.setWidths(width);
    table.setWidth(90); // 占页面宽度比例
    table.setAlignment(Element.ALIGN_CENTER); // 居中
    table.setAlignment(Element.ALIGN_MIDDLE); // 垂直居中
    table.setAutoFillEmptyCells(true); // 自动填满
    table.setBorderWidth(1); // 边框宽度
    // 设置表头
    Cell haderCell = new Cell("表格表头");
    haderCell.setHeader(true);
    haderCell.setColspan(3);
    table.addCell(haderCell);
    table.endHeaders();

    Font fontChinese = new Font(bfChinese, 12, Font.NORMAL, Color.GREEN);
    Cell cell1 = new Cell(new Paragraph("这是一个3*3测试表格数据", fontChinese));
    cell1.setVerticalAlignment(Element.ALIGN_MIDDLE);
    table.addCell(cell1);
    table.addCell(new Cell("#1"));
    table.addCell(new Cell("#2"));
    table.addCell(new Cell("#3"));
    try {
      document.add(title);
      document.add(context);
      document.add(title);
      document.add(jsdw);
      document.add(jsdwnr);
      document.add(tbnr);
      document.add(ztqk);

    } catch (DocumentException e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }
    //    字体文档样式
    // [接收单位内容]
    Font ccFont = DocStyleUtils.setFontStyle("仿 宋", 22f, Font.NORMAL);
    StringBuffer sbs = new StringBuffer();
    sbs.append("很早之前反正全部都看完了···\n" + "很多字打出来就被和谐。。。\n" + "内容 和4L 差不多的意思。");
    Paragraph wordFile = DocStyleUtils.setParagraphStyle(sbs.toString(), ccFont, 32f, 1.2f);

    StringBuffer sbs1 = new StringBuffer();
    sbs1.append("　东南西北皆有阵，唯有中心把人困，如若不从四面起，田中有井化自己。\n" + "　　各路鬼怪都来说，阴兵阴将入中国，华夏儿郎今不醒，侵略之人日建多。");
    Paragraph fs = DocStyleUtils.setParagraphStyle(sbs1.toString(), ccFont, 32f, 1.2f);

    document.add(wordFile);
    document.add(fs);
    // 表格----------------------------
    BaseFont bcFont =
        BaseFont.createFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);

    Font tableNameFont = new Font(bcFont, 16, Font.BOLD);
    Font cellHeadFont = new Font(bcFont, 12, Font.BOLD);
    Paragraph tablename1 =
        DocStyleUtils.setParagraphStyle("各单位线路故障情况统计", tableNameFont, 0.1f, Paragraph.ALIGN_CENTER);
    tablename1.setLeading(0, 0);
    tablename1.setSpacingAfter(0);
    Font cellFont = new Font(bcFont, 12, Font.NORMAL);
    try {
      // 创建一 表格(必须指定列，也可指定列和行)

      Table table1 = new Table(7);

      table1.setAlignment(Element.ALIGN_CENTER);
      table1.setWidth(100f);
      // table1.setBorderWidthTop(0);
      table1.setCellsFitPage(true);
      table1.setBorderWidthTop(0);
      table1.setTop(0);

      table1.setOffset(0);
      // 设置每列所占比例
      int[] withs = {3, 3, 3, 3, 3, 4, 3};
      try {
        table1.setWidths(withs);
      } catch (DocumentException e1) {
        // TODO Auto-generated catch block
        e1.printStackTrace();
      }
      // 表头
      Cell[] cellHeaders = new Cell[7];

      // RtfFont cellFont = new RtfFont("仿宋_GB2312",9,Font.NORMAL,Color.BLACK);

      cellHeaders[0] = new Cell(new Phrase("序号", cellHeadFont));
      cellHeaders[1] = new Cell(new Phrase("单位", cellHeadFont));
      cellHeaders[2] = new Cell(new Phrase("重合不良", cellHeadFont));
      cellHeaders[3] = new Cell(new Phrase("重合良好", cellHeadFont));
      cellHeaders[4] = new Cell(new Phrase("接地故障", cellHeadFont));
      cellHeaders[5] = new Cell(new Phrase("分支线停电", cellHeadFont));
      cellHeaders[6] = new Cell(new Phrase("合计", cellHeadFont));

      for (int i = 0; i <= 6; i++) {
        // 居中显示
        cellHeaders[i].setHorizontalAlignment(Element.ALIGN_CENTER);
        // 纵向居中显示
        cellHeaders[i].setVerticalAlignment(Element.ALIGN_MIDDLE);
        table1.addCell(cellHeaders[i]);
      }

      // 第二行合并跨越
      Cell cell001 = new Cell(new Phrase("tjdwmc提交单位名称", cellFont));
      cell001.setColspan(2);
      cell001.setHorizontalAlignment(Element.ALIGN_CENTER);
      cell001.setVerticalAlignment(Element.ALIGN_MIDDLE);
      table1.addCell(cell001);

      Cell[] cellFirstRow = new Cell[5];

      cellFirstRow[0] = new Cell(new Phrase(String.valueOf("chblSum"), cellFont));
      cellFirstRow[1] = new Cell(new Phrase(String.valueOf("chlhSum"), cellFont));
      cellFirstRow[2] = new Cell(new Phrase(String.valueOf("jdgzSum"), cellFont));
      cellFirstRow[3] = new Cell(new Phrase(String.valueOf("fzxtdSum"), cellFont));
      cellFirstRow[4] =
          new Cell(new Phrase(String.valueOf("chblSum+chlhSum+jdgzSum+fzxtdSum"), cellFont));

      for (int i = 0; i < cellFirstRow.length; i++) {
        // 居中显示
        cellFirstRow[i].setHorizontalAlignment(Element.ALIGN_CENTER);
        // 纵向居中显示
        cellFirstRow[i].setVerticalAlignment(Element.ALIGN_MIDDLE);
        table1.addCell(cellFirstRow[i]);
      }

      int i = 0;
      for (int j = 0; j < 5; j++) {
        i = i + 1;

        Cell[] cell = new Cell[7];
        cell[0] = new Cell(new Phrase(String.valueOf(i), cellFont));
        cell[1] = new Cell(new Phrase(String.valueOf("map.get()"), cellFont));
        cell[2] = new Cell(new Phrase(String.valueOf("map.get(  重合不良 )"), cellFont));
        cell[3] = new Cell(new Phrase(String.valueOf("所发生的"), cellFont));
        cell[4] = new Cell(new Phrase(String.valueOf("接地故障"), cellFont));
        cell[5] = new Cell(new Phrase(String.valueOf("分支线停电"), cellFont));
        cell[6] =
            new Cell(
                new Phrase(
                    String.valueOf(
                        Integer.valueOf(30)
                            + Integer.valueOf(10)
                            + Integer.valueOf(11)
                            + Integer.valueOf(11)),
                    cellFont));
        for (int k = 0; k < cell.length; k++) {
          // 居中显示
          cell[k].setHorizontalAlignment(Element.ALIGN_CENTER);
          // 纵向居中显示
          cell[k].setVerticalAlignment(Element.ALIGN_MIDDLE);
          table1.addCell(cell[k]);
        }
      }

      document.add(tablename1);
      document.add(table1);
      document.add(table);
      document.add(p1Content);
      document.close();
    } catch (DocumentException e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }
    // 导出Zip

  }

  @GetMapping("/exportzip")
  public void exportZip(HttpServletResponse response) {
    this.setResponseHeaderForWord(response, "zipName.zip", "iso-8859-1");
    ZipOutputStream out = null;
    ServletOutputStream sOut = null;
    try {
      sOut = response.getOutputStream();
      out = new ZipOutputStream(sOut);
      // 没有这一行中文文件名将会乱码
      out.setEncoding(System.getProperty("sun.jnu.encoding"));
      out.setEncoding("gbk");
      List<FileBean> fileList = new ArrayList<>();

      FileBean fileBean = new FileBean();
      FileBean fileBean1 = new FileBean();
      fileBean.setFileName("11teszdzt.docx");
      fileBean1.setFileName("1zt.docx");
      fileBean.setFilePath("D:/tmp/");
      fileBean1.setFilePath("D:/tmp/");
      fileList.add(fileBean);
      fileList.add(fileBean1);

      for (Iterator<FileBean> it = fileList.iterator(); it.hasNext(); ) {
        FileBean file = it.next();
        ZipUtils.doCompress(file.getFilePath() + file.getFileName(), out);
        response.flushBuffer();
      }
    } catch (IOException e1) {
      e1.printStackTrace();
    } finally {
      if (out != null) {
        try {
          out.close();
        } catch (IOException e) {
          e.printStackTrace();
        }
      }
      if (sOut != null) {
        try {
          sOut.close();
        } catch (IOException e) {
          e.printStackTrace();
        }
      }
    }
    System.out.println("第三个文件导出成功结束");
  }

  public void exportZiptest(HttpServletResponse response) {
    this.setResponseHeaderForWord(response, "zipName.zip", "iso-8859-1");
    ZipOutputStream out = null;
    ServletOutputStream sOut = null;
    try {
      sOut = response.getOutputStream();
      out = new ZipOutputStream(sOut);
      // 没有这一行中文文件名将会乱码
      out.setEncoding(System.getProperty("sun.jnu.encoding"));
      out.setEncoding("gbk");
      List<FileBean> fileList = new ArrayList<>();

      FileBean fileBean = new FileBean();
      FileBean fileBean1 = new FileBean();
      fileBean.setFileName("11teszdzt.docx");
      fileBean1.setFileName("1zt.docx");
      fileBean.setFilePath("D:/tmp");
      fileBean1.setFilePath("D:/tmp");
      fileList.add(fileBean);
      fileList.add(fileBean1);

      for (Iterator<FileBean> it = fileList.iterator(); it.hasNext(); ) {
        FileBean file = it.next();
        ZipUtils.doCompress(file.getFilePath() + file.getFileName(), out);
        response.flushBuffer();
      }
    } catch (IOException e1) {
      e1.printStackTrace();
    } finally {
      if (out != null) {
        try {
          out.close();
        } catch (IOException e) {
          e.printStackTrace();
        }
      }
      if (sOut != null) {
        try {
          sOut.close();
        } catch (IOException e) {
          e.printStackTrace();
        }
      }
    }
    System.out.println("第三个文件导出成功");
    // 在最后删除D:\temDictionary文件夹下的所有文件(注意要有两个反斜杠)
   /* String temDictionary = "";
    temDictionary = temDictionary.substring(0, temDictionary.lastIndexOf("/"));
    File dir = new File(temDictionary);
    removeDir(dir);*/
    System.out.println("第三个文件导出成功结束");
  }
  /**
   * 删除指定文件夹下所有文件
   *
   * @param dir
   */
  private void removeDir(File dir) {
    // TODO Auto-generated method stub
    File[] files = dir.listFiles();
    if (files != null) {
      for (File file : files) {
        if (!file.isDirectory()) {
          file.delete();
        } else {
          removeDir(file);
        }
      }
    }
  }

  public static void main(String[] args) {
    ZipController zc = new ZipController();
    String file = "D:\\tmp/1zt.docx";
    try {
      zc.createDocContext(file, "测试sdfdsfdfdsfsdfdsiText导出Word文档");
      /*
      List<FileBean> fileList = new ArrayList<>();
      FileBean fileBean = new FileBean();
      fileBean.setFileName("zt.docx");
      // fileBean.setFilePath(path + "WEB-INF" +
      // File.separator+ "temDictionary" + File.separator);
      fileBean.setFilePath("D:/tmp");
      fileList.add(fileBean);*/
    } catch (Exception e) {
      e.printStackTrace();
    }
  }
/**
 * @Function TODO
 * @author   不弹出下载框, 直接下载
 * @Date     2020/6/12 13:45
 * @param    
 * @return   
 */
  public void setResponseHeaderForWord(HttpServletResponse response, String fileName, String bm) {
    //    fileName = CodeSecurityUtil.validisOkFileName(fileName);
    try {
      response.setContentType("application/octet-stream;charset=UTF-8");
      response.setHeader(
          "Content-disposition",
          String.format("attachment; filename=\"%s\"", new String(fileName.getBytes(), bm)));

      response.addHeader("Pargam", "no-cache");
      response.addHeader("Cache-Control", "no-cache");
    } catch (Exception e) {
    }
  }
}
