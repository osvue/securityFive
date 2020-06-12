package org.osvue.formlogin.controller;

import java.io.IOException;
import java.net.URLEncoder;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.osvue.formlogin.utils.WorderToNewWordUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * @Author: Mr.Han @Description: securityFive @Date: Created in 2020/5/23_10:45 @Modified By: THE
 * GIFTED
 */
@Controller
public class HelloController {
  @GetMapping("/hello")
  public String hello() {
    return "login";
  }

  /**
   * @Function TODO
   *
   * @author 使用POI导出word @Date 2020/6/12 9:23
   * @param
   * @return
   */
  @GetMapping("/excel")
  public void excel(HttpServletRequest request, HttpServletResponse response) {
    Map m3 = new HashMap();
    m3.put("NAME", "Nihaoshijie ");
    m3.put("END_TIME", LocalDate.ofYearDay(2018, 10));
    m3.put("START_TIME", LocalDateTime.now());

    // 获取文件名称
    String reportname = (m3).get("NAME").toString();

    XWPFDocument document = new XWPFDocument();
    // 第一行大标题
    WorderToNewWordUtils.setContent(
        document, reportname, "宋体", 22, ParagraphAlignment.CENTER, true);
    // 第二行 中等标题
    WorderToNewWordUtils.setContent(
        document,
        "(" + m3.get("START_TIME").toString() + "-" + m3.get("END_TIME") + ")",
        "宋体",
        18,
        ParagraphAlignment.CENTER,
        false);
    Map<String, Object> map1 = new HashMap<String, Object>();
    // 生成表
    map1.put(
        "content",
        "默认生成的 access_token 其实就是一个 UUID 字符串。\n"
            + "getAccessTokenValiditySeconds 方法用来获取 access_token 的有效期，点进去这个方法，我们发现这个数字是从数据库中查询出来的，其实就是我们配置的 access_token 的有效期，我们配置的有效期单位是秒。\n"
            + "如果设置的 access_token 有效期大于 0，则调用 setExpiration 方法设置过期时间，过期时间就是在当前时间基础上加上用户设置的过期时间，注意乘以 1000 将时间单位转为毫秒。\n"
            + "接下来设置刷新 token 和授权范围 scope（刷新 token 的生成过程在 createRefreshToken 方法中，其实和 access_token 的生成过程类似）。\n"
            + "最后面 return 比较关键，这里会判断有没有 accessTokenEnhancer，如果 accessTokenEnhancer 不为 null，则在 accessTokenEnhancer 中再处理一遍才返回，accessTokenEnhancer 中再处理一遍就比较关键了，就是 access_token 转为 jwt 字符串的过程。\n"
            + "这里的 accessTokenEnhancer 实际上是一个 TokenEnhancerChain，这个链中有一个 delegates 变量保存了我们定义的两个 TokenEnhancer（auth-server 中定义的 JwtAccessTokenConverter 和 CustomAdditionalInformation），也就是说，我们的 access_token 信息将在这两个类中进行二次处理");

    map1.put("tabletilte", "全省时户数消耗情况");
    List<List<Map<String, Object>>> testList = new ArrayList<>();
    XWPFTable table =
        WorderToNewWordUtils.setTable(
            map1,
            document,
            new String[][] {
              new String[] {
                "序号", "单位", "等效总;用户数", "累计用户平均停电情况", "", "", "", "停电时户数消耗情况", "", "", ""
              },
              new String[] {
                "", "", "", "全口;径", "同比", "预;安排;停电", "故障;停电", "年度;目标", "累计;消耗", "同比", "本周;消耗"
              }
            },
            testList);

    WorderToNewWordUtils.mergeCellsHorizontal(table, 0, 3, 6);
    WorderToNewWordUtils.mergeCellsHorizontal(table, 0, 7, 10);
    //        WorderToNewWordUtils.mergeCellsHorizontal(table, 2, 0, 1);
    //
    //        WorderToNewWordUtils.mergeCellsVertically(table, 0, 0, 1);
    //        WorderToNewWordUtils.mergeCellsVertically(table, 1, 0, 1);
    //        WorderToNewWordUtils.mergeCellsVertically(table, 2, 0, 1);

    try {

      ServletOutputStream ServletOutputStream = response.getOutputStream();
      response.setContentType("application/octet-stream");
      response.setHeader(
          "Content-Disposition",
          "attachmen;filename=" + URLEncoder.encode(reportname + ".docx", "UTF-8"));
      document.write(ServletOutputStream);
      ServletOutputStream.flush();
      ServletOutputStream.close();
    } catch (IOException e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }
  }
}
