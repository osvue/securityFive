package org.osvue.formlogin.utils;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.regex.Pattern;

import javax.imageio.ImageIO;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFRun.FontCharRange;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

/**
 * 通过word模板生成新的word工具类
 *
 * @author zhiheng
 */
public class WorderToNewWordUtils {

  public static void setContent(
      XWPFDocument document,
      String text,
      String fontFamily,
      int fontSize,
      ParagraphAlignment paragraphAlignment,
      boolean bold) {
    XWPFParagraph paragraph = document.createParagraph();
    XWPFRun r = paragraph.createRun();
    r.setText(text);
    r.setFontFamily(fontFamily);

    r.setBold(bold);
    r.setFontSize(fontSize);

    paragraph.setAlignment(paragraphAlignment);
    paragraph.addRun(r);
  }

  public static void setTitle(XWPFDocument document, String title) {
    // 标题
    {
      XWPFParagraph paragraph = document.createParagraph();
      XWPFRun r = paragraph.createRun();
      r.setText(title);
      r.setFontFamily("宋体");

      r.setBold(true);
      r.setFontSize(16);

      paragraph.setAlignment(ParagraphAlignment.LEFT);
      paragraph.setFirstLineIndent(320);
      paragraph.addRun(r);
    }
  }

  public static XWPFTable setTable(
      Map<String, Object> map,
      XWPFDocument document,
      String[][] s,
      List<List<Map<String, Object>>> testList) {

    // 内容
    content(map, document);

    // 表格标题
    {
      XWPFParagraph paragraph = document.createParagraph();
      XWPFRun r = paragraph.createRun();
      r.setText(map.get("tabletilte").toString());
      r.setFontFamily("宋体");
      r.setFontSize(14);
      paragraph.setAlignment(ParagraphAlignment.CENTER);
      paragraph.setFirstLineIndent(2);
      paragraph.addRun(r);
    }

    XWPFTable table = document.createTable(s.length + testList.size(), s[0].length);
    CTTbl ttbl = table.getCTTbl();
    CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
    CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
    CTJc cTJc = tblPr.addNewJc();
    cTJc.setVal(STJc.Enum.forString("center"));
    tblWidth.setW(new BigInteger("8000"));
    tblWidth.setType(STTblWidth.DXA);

    // 表格首航
    {
      for (int i = 0; i < s.length; i++) {
        XWPFTableRow row = table.getRow(i);
        for (int j = 0; j < s[i].length; j++) {
          XWPFParagraph p = row.getCell(j).getParagraphs().get(0);
          XWPFRun r1 = p.createRun();
          table
              .getRow(i)
              .getCell(j)
              .getCTTc()
              .addNewTcPr()
              .addNewVAlign()
              .setVal(STVerticalJc.CENTER);
          table.getRow(i).getCell(j).setColor("CCCCCC");
          String[] tt = s[i][j].split(";");
          for (int k = 0; k < tt.length; k++) {
            r1.setText(tt[k]);
            r1.setBold(true);
            if (k != tt.length - 1) {
              r1.addBreak();
            }
          }

          r1.setFontFamily("宋体");
          p.setAlignment(ParagraphAlignment.CENTER);
        }
      }
    }

    // 表格
    {
      for (int i = 0; i < testList.size(); i++) {
        XWPFTableRow row = table.getRow(i + s.length);
        for (int j = 0; j < testList.get(i).size(); j++) {
          XWPFParagraph p = row.getCell(j).getParagraphs().get(0);
          XWPFRun r1 = p.createRun();
          table
              .getRow(i)
              .getCell(j)
              .getCTTc()
              .addNewTcPr()
              .addNewVAlign()
              .setVal(STVerticalJc.CENTER);
          r1.setText(testList.get(i).get(j).get("COLUMN").toString());
          r1.setFontFamily("宋体");
          if (testList.get(i).get(j).get("fill") != null) {
            row.getCell(j).setColor("ff0000");
          }

          p.setAlignment(ParagraphAlignment.CENTER);
        }
      }
    }

    return table;
  }

  public static XWPFTable setTable(
      Map<String, Object> map,
      XWPFDocument document,
      String[][] s,
      List<List<Map<String, Object>>> testList,
      String[] tts) {

    // 内容
    content(map, document);

    // 表格标题
    {
      XWPFParagraph paragraph = document.createParagraph();
      XWPFRun r = paragraph.createRun();
      r.setText(map.get("tabletilte").toString());
      r.setFontFamily("宋体");
      r.setFontSize(14);
      paragraph.setAlignment(ParagraphAlignment.CENTER);
      paragraph.setFirstLineIndent(2);
      paragraph.addRun(r);
    }

    XWPFTable table = document.createTable(s.length + testList.size(), s[0].length);
    CTTbl ttbl = table.getCTTbl();
    CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
    CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
    CTJc cTJc = tblPr.addNewJc();
    cTJc.setVal(STJc.Enum.forString("center"));
    tblWidth.setW(new BigInteger("8000"));
    tblWidth.setType(STTblWidth.DXA);

    // 表格首航
    {
      for (int i = 0; i < s.length; i++) {
        XWPFTableRow row = table.getRow(i);
        for (int j = 0; j < s[i].length; j++) {
          XWPFParagraph p = row.getCell(j).getParagraphs().get(0);
          XWPFRun r1 = p.createRun();
          table
              .getRow(i)
              .getCell(j)
              .getCTTc()
              .addNewTcPr()
              .addNewVAlign()
              .setVal(STVerticalJc.CENTER);
          table.getRow(i).getCell(j).setColor("CCCCCC");
          String[] tt = s[i][j].split(";");
          for (int k = 0; k < tt.length; k++) {
            r1.setText(tt[k]);
            r1.setBold(true);
            if (k != tt.length - 1) {
              r1.addBreak();
            }
          }

          r1.setFontFamily("宋体");
          p.setAlignment(ParagraphAlignment.CENTER);
        }
      }
    }

    // 表格
    {
      for (int i = 0; i < testList.size(); i++) {
        XWPFTableRow row = table.getRow(i + s.length);
        for (int j = 0; j < testList.get(i).size(); j++) {
          XWPFParagraph p = row.getCell(j).getParagraphs().get(0);
          XWPFRun r1 = p.createRun();
          table
              .getRow(i)
              .getCell(j)
              .getCTTc()
              .addNewTcPr()
              .addNewVAlign()
              .setVal(STVerticalJc.CENTER);
          r1.setText(testList.get(i).get(j).get("COLUMN").toString());
          r1.setFontFamily("宋体");
          if (testList.get(i).get(j).get("fill") != null) {
            row.getCell(j).setColor("ff0000");
          }

          p.setAlignment(ParagraphAlignment.CENTER);
        }
      }
    }

    // 表格标题 tt
    {
      for (String t : tts) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun r = paragraph.createRun();
        r.setText(t);
        r.setFontFamily("宋体");
        r.setFontSize(8);
        r.setColor("ff0000");
        // paragraph.setAlignment(ParagraphAlignment.Left);
        paragraph.setFirstLineIndent(420);
        // paragraph.addRun(r);
      }
    }

    return table;
  }

  private static void content(Map<String, Object> map, XWPFDocument document) {
    Pattern p = Pattern.compile("\\s*|\t|\r|\n");
    String s =
        p.matcher(map.get("content").toString())
            .replaceAll("")
            .replace("&nbsp;", "")
            .replace("&#39;", "'")
            .replace("&quot;", "\"")
            .replace("\\r", "")
            .replace("\\n", "")
            .replaceAll("<p>\\s*</p>", "")
            .replace("<br/>", "");

    XWPFParagraph paragraph = document.createParagraph();

    XWPFRun r2 = paragraph.createRun();
    r2.addTab(); // tab键
    r2.setBold(false); // 加粗
    //		runX.setCapitalized(false);//我也不知道这个属性做啥的
    // runX.setCharacterSpacing(5);//这个属性报错
    r2.setColor("BED4F1"); // 设置颜色--十六进制
    //		runX.setDoubleStrikethrough(false);//双删除线
    //		runX.setEmbossed(false);//浮雕字体----效果和印记（悬浮阴影）类似
    r2.setFontFamily("宋体"); // 字体
    //		runX.setFontFamily("华文新魏", FontCharRange.cs);//字体，范围----效果不详
    r2.setFontSize(14); // 字体大小
    //		runX.setImprinted(false);//印迹（悬浮阴影）---效果和浮雕类似
    //		runX.setItalic(false);//斜体（字体倾斜）
    // runX.setKerning(1);//字距调整----这个好像没有效果
    //		runX.setShadow(true);//阴影---稍微有点效果（阴影不明显）
    // runX.setSmallCaps(true);//小型股------效果不清楚
    // runX.setStrike(true);//单删除线（废弃）
    //		runX.setStrikeThrough(false);//单删除线（新的替换Strike）
    // runX.setSubscript(VerticalAlign.SUBSCRIPT);//下标(吧当前这个run变成下标)---枚举
    // runX.setTextPosition(20);//设置两行之间的行间距//runX.setUnderline(UnderlinePatterns.DASH_LONG);//各种类型的下划线（枚举）
    // runX0.addBreak();//类似换行的操作（html的  br标签）

    //		runX0.addCarriageReturn();//回车键

    // 注意：addTab()和addCarriageReturn() 对setText()的使用先后顺序有关：
    // 比如先执行addTab,再写Text这是对当前这个Text的Table，
    // 反之是对下一个run的Text的Tab效果
    if (s.indexOf("<p>") < 0) {
      r2.setText(s, 0);
    } else {
      r2.setText("", 0);
    }
    int i = 0;
    //      单独处理段落标签
    while (true) {
      int i1 = s.indexOf("<p>");
      int i2 = s.indexOf("</p>");
      if (i1 < 0) {
        break;
      }

      XmlCursor xmlCursor = paragraph.getCTP().newCursor();
      if (i != 0) {
        xmlCursor.toNextSibling();
      }
      paragraph = paragraph.getDocument().insertNewParagraph(xmlCursor);
      i++;
      //

      paragraph.setFirstLineIndent(420);
      String s1 = s.substring(i1, i2 + 4);

      s = s.substring(i2 + 4, s.length());

      s1 = s1.replace("<p>", "");
      s1 = s1.replace("</p>", "");

      while (true) {

        int i3 = s1.indexOf("<strong>");
        if (i3 < 0) {
          XWPFRun run = paragraph.createRun();
          run.setBold(false);
          run.setFontSize(14);
          run.setText(s1);
          break;
        }
        String s2 = s1.substring(0, i3);
        XWPFRun run = paragraph.createRun();
        run.setBold(true);
        run.setFontSize(14);
        run.setText(s2.replace("<strong>", "").replace("</strong>", ""));

        int i4 = s1.indexOf("</strong>");

        s2 = s1.substring(i3, i4 + 9);

        XWPFRun run2 = paragraph.createRun();
        run2.setBold(true);
        run2.setText(s2.replace("<strong>", "").replace("</strong>", ""));

        s1 = s1.substring(i4 + 9, s1.length());
      }
    }
  }

  public static void main2(String[] args) throws IOException {
    XWPFDocument document = new XWPFDocument();
    InputStream is = null;

    String[][] s =
        new String[][] {
          new String[] {
            "safsafsalkfdjasfjlksdfjalafjskfasj;fad;jfs;fjajfdas;jfasjfjfds;afjafjdafsdjdsfjdfssdsajdfjasdjffdjs",
            "1-2",
            "1-3"
          },
          new String[] {"2-1", "2-2", "2-3"},
        };

    XWPFTable table = document.createTable();

    for (int i = 0; i < s.length; i++) {
      XWPFTableRow row;
      if (i == 0) {
        row = table.getRow(0);
        for (int j = 0; j < s[i].length; j++) {
          XWPFTableCell cell;
          if (j == 0) {
            cell = row.getCell(0);
          } else {
            cell = row.addNewTableCell();
          }
          XWPFParagraph p = cell.addParagraph();
          XWPFRun r1 = p.createRun();
          cell.getCTTc().addNewTcPr().addNewVAlign().setVal(STVerticalJc.CENTER);

          String[] tt = s[i][j].split(";");
          for (int k = 0; k < tt.length; k++) {
            r1.setText(tt[k]);
            r1.addBreak();
          }

          r1.setFontFamily("宋体");
          p.setAlignment(ParagraphAlignment.CENTER);

          CTTc cttc = cell.getCTTc();
          CTTcPr cellPr = cttc.addNewTcPr();
          CTTblWidth cellw = cellPr.addNewTcW();
          cellw.setType(STTblWidth.NIL);
          cellw.setW(BigInteger.valueOf(100));

          cell.setColor("CCCCCC");

          row.setHeight(380);

          // 边线颜色
          CTTblBorders borders = table.getCTTbl().getTblPr().addNewTblBorders();
          CTBorder hBorder = borders.addNewInsideH();
          hBorder.setVal(STBorder.Enum.forString("dashed"));
          hBorder.setSz(new BigInteger("1"));
          hBorder.setColor("0000FF");
        }
      } else {
        row = table.createRow();
        for (int j = 0; j < s[i].length; j++) {
          XWPFTableCell cell;
          cell = row.getCell(j);
          XWPFParagraph p = cell.addParagraph();
          XWPFRun r1 = p.createRun();
          cell.getCTTc().addNewTcPr().addNewVAlign().setVal(STVerticalJc.CENTER);
          ;

          CTTc cttc = cell.getCTTc();
          CTTcPr cellPr = cttc.addNewTcPr();
          cellPr.addNewTcW().setW(BigInteger.valueOf(1255));

          r1.setText(s[i][j]);
          r1.setFontFamily("宋体");
          p.setAlignment(ParagraphAlignment.CENTER);
        }
      }
    }

    XWPFParagraph paragraph = document.createParagraph();

    XWPFRun r = paragraph.createRun();
    // cell.getCTTc().addNewTcPr().addNewVAlign().setVal(STVerticalJc.CENTER);;
    r.setText("ssaolkjjl");
    r.setFontFamily("宋体");
    paragraph.setAlignment(ParagraphAlignment.CENTER);

    paragraph.addRun(r);

    XWPFTable table1 = document.createTable(5, 5);

    CTTbl ttbl = table1.getCTTbl();
    CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
    CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
    CTJc cTJc = tblPr.addNewJc();
    cTJc.setVal(STJc.Enum.forString("center"));
    tblWidth.setW(new BigInteger("8000"));
    tblWidth.setType(STTblWidth.DXA);
    XWPFParagraph p = table1.getRow(0).getCell(0).addParagraph();
    XWPFRun r1 = p.createRun();
    table1.getRow(0).getCell(0).getCTTc().addNewTcPr().addNewVAlign().setVal(STVerticalJc.CENTER);
    ;
    r1.setText("1255wdfjskdfjlkasdflkasjflkasjfjaskdfjlsdflaslfasljdfl;");
    r1.addBreak();
    r1.addBreak();
    r1.setFontFamily("宋体");
    r1.addBreak();
    p.setAlignment(ParagraphAlignment.CENTER);

    p.setFirstLineIndent(2);
    XWPFParagraph p2 = table1.getRow(0).getCell(1).addParagraph();
    XWPFRun r21 = p2.createRun();
    table1.getRow(0).getCell(1).getCTTc().addNewTcPr().addNewVAlign().setVal(STVerticalJc.CENTER);
    ;
    r21.setText("1255wdfjskdfjlkasdflkasjflkasjfjaskdfjlsdflaslfasljdfl;");
    r21.addBreak();
    r21.addBreak();
    r21.setFontFamily("宋体");
    r21.addBreak();
    p2.setAlignment(ParagraphAlignment.CENTER);

    CTTcPr tcpr = table1.getRow(0).getCell(0).getCTTc().addNewTcPr();

    CTTblWidth cellw = tcpr.addNewTcW();

    cellw.setType(STTblWidth.DXA);

    cellw.setW(BigInteger.valueOf(5));

    /*
     * mergeCellsHorizontal(table1, 0, 0, 1); mergeCellsHorizontal(table1, 1, 0, 1);
     * mergeCellsVertically(table1, 0, 0, 1);
     */

    /**
     * XWPFTableRow row = table.createRow(); XWPFTableCell cell = row.addNewTableCell();
     *
     * <p>XWPFParagraph p = cell.addParagraph(); XWPFRun r1 = p.createRun(); //
     * cell.getCTTc().addNewTcPr().addNewVAlign().setVal(STVerticalJc.CENTER);;
     * r1.setText("12312424124"); r1.setFontFamily("宋体"); p.setAlignment(ParagraphAlignment.CENTER);
     */
    FileOutputStream out =
        new FileOutputStream(
            "C:\\Users\\Administrator\\Desktop\\1.doc"); // true为进入追加模式，false为覆盖原有内容
    document.write(out);
  }

  /**
   * 设置行合并属性
   *
   * @author Christmas_G
   * @date 2019-05-31 14:08:02
   * @param tableCell
   * @param mergeVlaue
   */
  private static void setRowMerge(XWPFTableCell tableCell, STMerge.Enum mergeVlaue) {
    CTTc ctTc = tableCell.getCTTc();
    CTTcPr cpr = ctTc.isSetTcPr() ? ctTc.getTcPr() : ctTc.addNewTcPr();
    CTVMerge merge = cpr.isSetVMerge() ? cpr.getVMerge() : cpr.addNewVMerge();
    merge.setVal(mergeVlaue);
  }

  /**
   * 设置列合并属性
   *
   * @author Christmas_G
   * @date 2019-05-31 14:07:50
   * @param tableCell
   * @param mergeVlaue
   */
  private static void setColMerge(XWPFTableCell tableCell, STMerge.Enum mergeVlaue) {
    CTTc ctTc = tableCell.getCTTc();
    CTTcPr cpr = ctTc.isSetTcPr() ? ctTc.getTcPr() : ctTc.addNewTcPr();
    CTHMerge merge = cpr.isSetHMerge() ? cpr.getHMerge() : cpr.addNewHMerge();
    merge.setVal(mergeVlaue);
  }

  /** @Description: 跨列合并 */
  public static void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
    for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
      XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
      if (cellIndex == fromCell) {
        // The first merged cell is set with RESTART merge value
        cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
      } else {
        // Cells which join (merge) the first one, are set with CONTINUE
        cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
      }
    }
  }
  /**
   * 根据模板生成新word文档 判断表格是需要替换还是需要插入，判断逻辑有$为替换，表格无$为插入
   *
   * @param inputUrl 模板存放地址
   * @param *outPutUrl 新文档存放地址
   * @param textMap 需要替换的信息集合
   * @param tableList 需要插入的表格信息集合
   * @return 成功返回true,失败返回false
   */
  public static XWPFDocument changWord(
      String inputUrl,
      Map<String, String> textMap,
      List<List<List<Map<String, Object>>>> tableList,
      Map<String, String> pic) {

    // 模板转换默认成功
    boolean changeFlag = true;
    XWPFDocument document = null;
    try {
      // 获取docx解析对象
      document = new XWPFDocument(POIXMLDocument.openPackage(inputUrl));
      // 解析替换文本段落对象
      WorderToNewWordUtils.changeText(document, textMap);
      // 解析替换表格对象
      WorderToNewWordUtils.changeTable(document, textMap, tableList);
      // 解析替换文本段落对象
      WorderToNewWordUtils.changePic(document, pic);

      // 获取本地图片

      /*            //生成新的word
      File file = new File(outputUrl);
        FileOutputStream stream = new FileOutputStream(file);
        document.write(stream);
        stream.close();*/

    } catch (IOException e) {
      e.printStackTrace();
      changeFlag = false;
    }
    return document;
    // return changeFlag;

  }

  /**
   * 替换段落文本
   *
   * @param document docx解析对象
   * @param textMap 需要替换的信息集合
   */
  public static void changeText(XWPFDocument document, Map<String, String> textMap) {
    // 获取段落集合
    List<XWPFParagraph> paragraphs = document.getParagraphs();

    List<XWPFParagraph> _paragraphs = new ArrayList<XWPFParagraph>();
    for (XWPFParagraph xw : paragraphs) {
      _paragraphs.add(xw);
    }

    paragraphs.size();
    for (XWPFParagraph paragraph : _paragraphs) {
      // 判断此段落时候需要进行替换
      String text = paragraph.getText();
      if (checkText(text)) {
        List<XWPFRun> runs = paragraph.getRuns();
        // 替换模板原来位置
        changeValue(runs.get(0).toString(), textMap, paragraph, runs.get(0));
      }
    }
  }

  /**
   * 替换段落文本
   *
   * @param document docx解析对象
   * @param textMap 需要替换的信息集合
   */
  public static void changePic(XWPFDocument document, Map<String, String> textMap) {
    // 获取段落集合
    List<XWPFParagraph> paragraphs = document.getParagraphs();

    List<XWPFParagraph> _paragraphs = new ArrayList<XWPFParagraph>();
    for (XWPFParagraph xw : paragraphs) {
      _paragraphs.add(xw);
    }

    paragraphs.size();
    for (XWPFParagraph paragraph : _paragraphs) {
      // 判断此段落时候需要进行替换
      String text = paragraph.getText();
      if (checkText2(text)) {
        List<XWPFRun> runs = paragraph.getRuns();
        // 替换模板原来位置
        changePic(runs.get(0).toString(), textMap, paragraph, runs.get(0));
      }
    }
  }

  /**
   * 替换表格对象方法
   *
   * @param document docx解析对象
   * @param textMap 需要替换的信息集合
   * @param tableList 需要插入的表格信息集合
   */
  public static void changeTable(
      XWPFDocument document,
      Map<String, String> textMap,
      List<List<List<Map<String, Object>>>> tableList) {
    // 获取表格对象集合
    List<XWPFTable> tables = document.getTables();
    for (int i = 0; i < tables.size(); i++) {
      // 只处理行数大于等于2的表格，且不循环表头
      XWPFTable table = tables.get(i);
      // if(table.getRows().size()>1){
      // 判断表格是需要替换还是需要插入，判断逻辑有$为替换，表格无$为插入
      /* if(checkText(table.getText())){
          List<XWPFTableRow> rows = table.getRows();
          //遍历表格,并替换模板
          eachTable(rows, textMap);
      }else{*/
      insertTable(table, tableList.get(i), document);
      /*                }*/
      // }
    }
  }

  /*    */
  /**
   * 遍历表格
   *
   * @param rows 表格行对象
   * @param textMap 需要替换的信息集合
   */
  /*
  public static void eachTable(List<XWPFTableRow> rows ,Map<String, String> textMap){
      for (XWPFTableRow row : rows) {
          List<XWPFTableCell> cells = row.getTableCells();
          for (XWPFTableCell cell : cells) {
              //判断单元格是否需要替换
              if(checkText(cell.getText())){
                  List<XWPFParagraph> paragraphs = cell.getParagraphs();
                  for (XWPFParagraph paragraph : paragraphs) {
                      List<XWPFRun> runs = paragraph.getRuns();
                      for (XWPFRun run : runs) {
                          run.setText(changeValue(run.toString(), textMap),0);
                      }
                  }
              }
          }
      }
  }*/

  /**
   * 为表格插入数据，行数不够添加新行
   *
   * @param table 需要插入数据的表格
   * @param tableList 插入数据集合
   */
  public static void insertTable(
      XWPFTable table, List<List<Map<String, Object>>> tableList, XWPFDocument document) {
    // 创建行,根据需要插入的数据添加新行，不处理表头
    /* for(int i = 1; i < tableList.size(); i++){
        XWPFTableRow row =table.createRow();
    }*/
    // 遍历表格插入数据
    List<XWPFTableRow> rows = table.getRows();

    if (rows.size() > 15) {
      for (int i = rows.size() - 15; i < rows.size(); i++) {
        XWPFTableRow newRow = table.getRow(i);
        List<XWPFTableCell> cells = newRow.getTableCells();
        for (int j = 0; j < cells.size(); j++) {
          XWPFTableCell cell = cells.get(j);
          System.out.println(tableList);
          System.out.println(i);
          System.out.println(rows.size() - 15);
          System.out.println(j);
          if (tableList.get(i - (rows.size() - 15)).get(j).get("fill") != null) {
            cell.setColor("ff0000");
          }
          XWPFParagraph p = cell.getParagraphs().get(0);
          XWPFRun r = p.createRun();
          // cell.getCTTc().addNewTcPr().addNewVAlign().setVal(STVerticalJc.CENTER);;
          r.setText(tableList.get(i - (rows.size() - 15)).get(j).get("COLUMN").toString());
          r.setFontFamily("宋体");
          p.setAlignment(ParagraphAlignment.CENTER);
          // cell.setParagraph(p);
          // cell.setText(tableList.get(i-1)[j]);
        }
      }
    } else {

      for (int i = 0; i < tableList.size(); i++) {
        XWPFTableRow newRow = table.createRow();
        List<XWPFTableCell> cells = newRow.getTableCells();
        for (int j = 0; j < tableList.get(i).size(); j++) {
          XWPFTableCell cell = cells.get(j);
          XWPFParagraph p = cell.getParagraphs().get(0);
          XWPFRun r = p.createRun();
          // cell.getCTTc().addNewTcPr().addNewVAlign().setVal(STVerticalJc.CENTER);;
          r.setText(tableList.get(i).get(j).get("COLUMN").toString());
          r.setFontFamily("宋体");
          p.setAlignment(ParagraphAlignment.CENTER);
        }
      }
    }
  }

  /**
   * 判断文本中时候包含$
   *
   * @param text 文本
   * @return 包含返回true,不包含返回false
   */
  public static boolean checkText(String text) {
    boolean check = false;
    if (text.indexOf("$") != -1) {
      check = true;
    }
    return check;
  }

  /**
   * 匹配传入信息集合与模板
   *
   * @param value 模板需要替换的区域
   * @param textMap 传入信息集合
   * @return 模板需要替换区域信息集合对应值
   */
  public static String changeValue(
      String value, Map<String, String> textMap, XWPFParagraph paragraph, XWPFRun rsdfun) {
    Set<Entry<String, String>> textSets = textMap.entrySet();
    for (Entry<String, String> textSet : textSets) {
      // 匹配模板与替换值 格式${key}
      String key = "${" + textSet.getKey() + "}";
      Pattern p = Pattern.compile("\\s*|\t|\r|\n");
      String s =
          p.matcher(value.replace(key, textSet.getValue()))
              .replaceAll("")
              .replace("&nbsp;", "")
              .replace("&#39;", "'")
              .replace("&quot;", "\"")
              .replace("\\r", "")
              .replace("\\n", "")
              .replaceAll("<p>\\s*</p>", "")
              .replace("<br/>", "");

      if (s.indexOf("<p>") < 0 && s.indexOf("${") < 0) {
        if ("${TITLE}".equals(key)) {
          rsdfun.setText("", 0);
          XWPFRun run = paragraph.createRun();
          run.setFontFamily("宋体");
          run.setFontSize(30);
          run.setText(s, 0);
          run.setBold(true);
        } else {
          rsdfun.setText("", 0);
          XWPFRun run = paragraph.createRun();
          run.setFontFamily("宋体");
          run.setFontSize(16);
          run.setText(s, 0);
        }
      } else {
        rsdfun.setText("", 0);
      }
      int i = 0;
      while (true) {
        int i1 = s.indexOf("<p>");
        int i2 = s.indexOf("</p>");
        if (i1 < 0) {
          break;
        }

        XmlCursor xmlCursor = paragraph.getCTP().newCursor();
        if (i != 0) {
          xmlCursor.toNextSibling();
          paragraph = paragraph.getDocument().insertNewParagraph(xmlCursor);
        }
        i++;
        //

        paragraph.setFirstLineIndent(420);
        String s1 = s.substring(i1, i2 + 4);

        s = s.substring(i2 + 4, s.length());

        s1 = s1.replace("<p>", "");
        s1 = s1.replace("</p>", "");
        while (true) {

          int i3 = s1.indexOf("<strong>");
          if (i3 < 0) {
            XWPFRun run = paragraph.createRun();
            run.setBold(false);

            if ("${Q}".equals(key)) {
              run.setBold(true);
              run.setFontSize(16);
            }

            run.setText(HTMLSpirit.delHTMLTag(s1));
            break;
          }
          String s2 = s1.substring(0, i3);
          XWPFRun run = paragraph.createRun();
          run.setBold(false);
          run.setText(HTMLSpirit.delHTMLTag(s2.replace("<strong>", "").replace("</strong>", "")));

          int i4 = s1.indexOf("</strong>");

          s2 = s1.substring(i3, i4 + 9);

          XWPFRun run2 = paragraph.createRun();
          run2.setBold(true);
          run2.setText(HTMLSpirit.delHTMLTag(s2.replace("<strong>", "").replace("</strong>", "")));

          s1 = s1.substring(i4 + 9, s1.length());
        }
      }
    }

    return value;
  }

  public static String changePic(
      String value, Map<String, String> textMap, XWPFParagraph paragraph, XWPFRun rsdfun) {
    Set<Entry<String, String>> textSets = textMap.entrySet();
    for (Entry<String, String> textSet : textSets) {
      // 匹配模板与替换值 格式${key}
      String key = "#{" + textSet.getKey() + "}";

      ;
      /*
       * Pattern p = Pattern.compile("\\s*|\t|\r|\n"); String s =
       * p.matcher(value.replace(key,
       * textSet.getValue())).replaceAll("").replace("&nbsp;", "").replace("&#39;",
       * "'").replace("&quot;", "\"") .replace("\\r", "").replace("\\n",
       * "").replaceAll("<p>\\s*</p>", "").replace("<br/>", "");
       */

      rsdfun.setText("", 0);
      XWPFRun run = paragraph.createRun();

      File file = new File(value.replace(key, textSet.getValue()));
      InputStream in;
      BufferedImage image;
      try {
        in = new FileInputStream(file);
        image = ImageIO.read(file);

        run.addPicture(
            in,
            org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_JPEG,
            "",
            Units.pixelToEMU(600),
            Units.pixelToEMU(300));
      } catch (InvalidFormatException | IOException e) {
        // TODO Auto-generated catch block
        e.printStackTrace();
      }
    }

    return value;
  }
  /**
   * @Description: 跨行合并
   *
   * @see -http://stackoverflow.com/questions/24907541/row-span-with-xwpftable
   */
  public static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
    for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
      XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
      if (rowIndex == fromRow) {
        // The first merged cell is set with RESTART merge value
        cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
      } else {
        // Cells which join (merge) the first one, are set with CONTINUE
        cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
      }
    }
  }

  public static boolean checkText2(String text) {
    boolean check = false;
    if (text.indexOf("#") != -1) {
      check = true;
    }
    return check;
  }
}
