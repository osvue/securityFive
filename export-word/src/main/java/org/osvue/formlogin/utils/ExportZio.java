package org.osvue.formlogin.utils;

import javax.servlet.http.HttpServletResponse;

/**
 * @Author: Mr.Han
 * @Description: securityFive
 * @Date: Created in 2020/6/11_16:54
 * @Modified By: THE GIFTED
 */
public class ExportZio {


  public void setResponseHeaderForWord(HttpServletResponse response,
      String fileName, String bm) {
//    验证文件名是否合法 判断文件名称中是否存在特殊字符
//    fileName = CodeSecurityUtil.validisOkFileName(fileName);
    try {
      response.setContentType("application/octet-stream;charset=UTF-8");
      response.setHeader("Content-disposition", String.format(
          "attachment; filename=\"%s\"",
          new String(fileName.getBytes(), bm)));

      response.addHeader("Pargam", "no-cache");
      response.addHeader("Cache-Control", "no-cache");
    } catch (Exception e) {
    }
  }
}
 