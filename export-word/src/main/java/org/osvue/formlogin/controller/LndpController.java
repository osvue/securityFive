
package org.osvue.formlogin.controller;

/**
 * @Author: Mr.Han
 * @Description: securityFive
 * @Date: Created in 2020/10/29_10:18
 * @Modified By: THE GIFTED
 */

public class LndpController {
}
/*
  @Autowired(required = false)
  public SqlSessionTemplate sqlsessionTemplate;

  @Autowired
  JdbcTemplate jdbcTemplate;

  private static final String[] table_names = new String[] { "YYFX_PWTD", "YYFX_PFTD", "YYFX_ZGZXL", "YYFX_YCSB",
      "YYFX_PBYC", "YYFX_XQJH", "YYFX_CZPZX", "YYFX_XSB", "YYFX_QXQK", "YYFX_TDBS", "YYFX_FWZB", "YYFX_JKQK",
      "YYFX_YKQLC", "YYFX_YKBZ", "YYFX_DYYK", "YYFX_YPD", "YYFX_XTY", "YYFX_ZGZMX", "YYFX_PFTDMX", "YYYY_JHYS_YY",
      "YYYY_JHZX_YY", "YYYY_CZP_YY"

      , "REPORT_TDSHS", "REPORT_TDYCQK", "REPORT_PWGZQK", "REPORT_DDYSXBPH", "REPORT_ZGZQK", "REPORT_PWBXQK",
      "REPORT_PWBXTSQK", "REPORT_GFSQK", "REPORT_LCJYSGZTD", "REPORT_BXSYSGZTD" };

  private static final String[] table_cn_names = new String[] { "配网停电情况", "频繁停电情况", "线路负载监控", "异常设备监测", "配变异常监测预警情况",
      "全省巡视、消缺计划执行情况", "操作票执行", "全省新投运设备数量", "全省故障抢修服务情况", "全省停电信息报送情况", "95598 服务指标情况", "全省营业厅监控情况",
      "10 千伏业扩全流程时长管控情况", "业扩报装全流程管控情况", "全省低压业扩报装接入容量情况", "营配调数据治理情况", "新投运设备明细（城网）", "全省重过载线路明细", "全省频繁停电线路明细",
      "计划延迟原因", "计划执行原因", "操作票原因"

      , "全省停电时户数消耗情况", "全省停电异常情况", "全省配网故障情况", "全省低电压、三相不平衡情况", "全省重过载情况", "全省配网报修情况", "全省配网报修投诉情况", "全省高负损情况",
      "两小时重复", "八小时" };

  private static final String[] table_names2 = new String[] { "JJRBG_GZQK", "JJRBG_TDYY", "JJRBG_TDBW",
      "JJRBG_PBYXQK", "JJRBG_GZBX", "JJRBG_DYBX", "JJRBG_BPBJ", "JJRBG_JRBD", "JJRBG_FB_PDGZ", "JJRBG_QXCLSC",
      "JJRBG_PWTS", "JJRBG_SDLR", "JJRBG_TDSHS", "JJRBG_YQFKGDBZQK" };
  private static final String[] table_cn_names2 = new String[] { "各单位线路故障情况统计", "各单位停电原因情况统计", "各单位停电部位情况统计",
      "配变运行情况 ", "各单位故障报修统计表", "各单位低压报修统计表", "各单位备品备件准备情况", "节日保电隐患治理情况统计", "配变过载原因及采取措施详表", "平均处理时长", "配网投诉",
      "手动录入", "停电时户数", "疫情防控供电保障情况" };

  private static final String[][] first_name = new String[][] { {}, {}, {} };


  */
/** 下载word *//*
http://127.0.0.1:8888/lndp/yyfx/saveReport/541?bbb_0=29&bbb_1=27
  @SuppressWarnings({ "rawtypes" })
  @RequestMapping("/saveReport/{zid}")
  public void saveReport(HttpServletRequest request, HttpServletResponse response, @PathVariable int zid,
      String ids) {

    request.setAttribute("role_city", "getUser(request).getCity()");

    Map<String, Object> obj = new HashMap<>();

    Map<String, Object> map = new HashMap<>();

    map.put("value",
        "select NAME,ID,CANCHANGE,TO_NUMBER(to_char(START_TIME,'MM'))||' 月 '||TO_NUMBER(to_char(START_TIME,'dd'))||' 日' as START_TIME,TO_NUMBER(to_char(END_TIME,'MM'))||' 月 '||TO_NUMBER(to_char(END_TIME,'dd'))||' 日'as END_TIME from YYFX_Z where id = "
            + zid);

    Map m3 = (Map) sqlsessionTemplate.selectOne("cn.wykj91.ppt.dao.BaseMapper.getList", map);
    String reportname = (m3).get("NAME").toString();

    XWPFDocument document = new XWPFDocument();
    WorderToNewWordUtils.setContent(document, reportname, "宋体", 22, ParagraphAlignment.CENTER, true);
    WorderToNewWordUtils.setContent(document,
        "(" + m3.get("START_TIME").toString() + "-" + m3.get("END_TIME") + ")", "宋体", 18,
        ParagraphAlignment.CENTER, false);
    for (String table_name : table_names) {
      map.put("value", "select * from user_col_comments where table_name='" + table_name + "'");
      List<Map<String, Object>> list = sqlsessionTemplate.selectList("cn.wykj91.ppt.dao.BaseMapper.getList", map);
      String sql = "select ";// sum(decode(ms,null,0,ms))ms,
      if (!"YYFX_XTY".equals(table_name) && !"YYFX_ZGZMX".equals(table_name) && !"YYFX_PFTDMX".equals(table_name)
          && !"YYYY_JHYS_YY".equals(table_name) && !"YYYY_JHZX_YY".equals(table_name)
          && !"YYYY_CZP_YY".equals(table_name) && !"REPORT_LCJYSGZTD".equals(table_name)
          && !"REPORT_BXSYSGZTD".equals(table_name)

      ) {
        for (Map<String, Object> m : list) {
          if (!"ID".equals(m.get("COLUMN_NAME")) && !"Z_ID".equals(m.get("COLUMN_NAME"))
              && !"CITY_ID".equals(m.get("COLUMN_NAME"))) {
            sql += " sum(decode(" + m.get("COLUMN_NAME") + ",NULL,0," + m.get("COLUMN_NAME") + "))"
                + m.get("COLUMN_NAME") + " ,";
          }
        }
        sql = sql.substring(0, sql.length() - 1);
        sql += " from " + table_name + " where z_id = " + zid;
        map.put("value", sql);
        obj.put(table_name + "_sum", sqlsessionTemplate.selectOne("cn.wykj91.ppt.dao.BaseMapper.getList", map));

        sql = "select ";// sum(decode(ms,null,0,ms))ms,
        for (Map<String, Object> m : list) {
          if (!"ID".equals(m.get("COLUMN_NAME")) && !"Z_ID".equals(m.get("COLUMN_NAME"))
              && !"CITY_ID".equals(m.get("COLUMN_NAME"))) {
            sql += " round(avg(" + m.get("COLUMN_NAME") + "),2)" + m.get("COLUMN_NAME") + " ,";
          }
        }
        sql = sql.substring(0, sql.length() - 1);
        sql += " from " + table_name + " where z_id = " + zid;
        map.put("value", sql);
        obj.put(table_name + "_avg", sqlsessionTemplate.selectOne("cn.wykj91.ppt.dao.BaseMapper.getList", map));
      }

      sql = "select b.NAME,c.CANCHANGE, ";
      for (Map<String, Object> m : list) {
        if (!"ID".equals(m.get("COLUMN_NAME")) && !"Z_ID".equals(m.get("COLUMN_NAME"))
            && !"CITY_ID".equals(m.get("COLUMN_NAME"))) {
          sql += " decode(" + m.get("COLUMN_NAME") + ",NULL,'—'," + m.get("COLUMN_NAME") + ")"
              + m.get("COLUMN_NAME") + " ,";
        }
      }
      sql = sql.substring(0, sql.length() - 1);
      sql += " from (select * from " + table_name + " where z_id='" + zid
          + (getUser(request).getCity() == null ? "'" : "' and city_id =" + getUser(request).getCity())
          + ")a left join ls_city b on a.city_id = b.id "
          + " left join YYFX_Z c on a.z_id = c.id order by b.id";

      map.put("value", sql);

      List<Map<String, Object>> list2 = sqlsessionTemplate.selectList("cn.wykj91.ppt.dao.BaseMapper.getList",
          map);
      obj.put(table_name, JSON.toJSON(list2));

      if ("YYFX_PFTD".equals(table_name)) {
        obj.put(table_name + "1", sort(list2, "PFTD", 4));
        obj.put(table_name + "2", sort(list2, "FIFTD", 2));
        obj.put(table_name + "3", sort(list2, "CFTD", 3));
        obj.put(table_name + "4", sort(list2, "GZTD", 4));
        obj.put(table_name + "5", sort(list2, "ZCTD", 4));
      }

      if ("YYFX_ZGZXL".equals(table_name)) {
        obj.put(table_name + "1", sort(list2, "ZGZXL", 3));
      }

      if ("YYFX_YCSB".equals(table_name)) {
        obj.put(table_name + "1", sort(list2, "YCPB", 3));
      }
      if ("YYFX_QXQK".equals(table_name)) {
        obj.put(table_name + "1", sort(list2, "XFSC", 3));
        obj.put(table_name + "2", sortDesc(list2, "JDL", 3));
      }

      if ("YYFX_FWZB".equals(table_name)) {
        obj.put(table_name + "1", sort(list2, "YXLKH", 3));
        obj.put(table_name + "2", sort(list2, "SCLKH", 3));
      }
    }

    String sql = "select * from yyfx_json where z_id = " + zid + " order by a";
    map.put("value", sql);
    List<Map<String, Object>> list2 = sqlsessionTemplate.selectList("cn.wykj91.ppt.dao.BaseMapper.getList", map);
    String[] name = new String[] { "一", "二", "三", "四", "五", "六", "七", "八", "九", "十" };
    int name_index = 0;

    String ZDTDCS_MAX = "select max(to_number(ZDTDCS)) as ZDTDCS_MAX from REPORT_TDYCQK where  z_id = " + zid;
    map.put("value", ZDTDCS_MAX);
    obj.put("ZDTDCS", sqlsessionTemplate.selectOne("cn.wykj91.ppt.dao.BaseMapper.getList", map));

    // 同比、环比
    String REPORT_TB_HB = "SELECT * FROM REPORT_S WHERE z_id = " + zid;
    map.put("value", REPORT_TB_HB);
    obj.put("REPORT_TB_HB", sqlsessionTemplate.selectOne("cn.wykj91.ppt.dao.BaseMapper.getList", map));

    */
/*
     * 000212151 String [] idss = ids.split(","); boolean b1_1 = false; boolean b1_2
     * = false; boolean b2_1 = false; boolean b2_2 = false;
     *//*


    Enumeration em = request.getParameterNames();
    List<String> l = new ArrayList<String>();
    while (em.hasMoreElements()) {
      String names = em.nextElement().toString();
      if (names.indexOf("aaa") >= 0 || names.indexOf("bbb") >= 0) {
        l.add(names);
      }
    }

//		Collections.sort(l, new Comparator<String>() {
//            @Override
//            public int compare(String o1, String o2) {
//                return Integer.parseInt(o1.replace("aaa_", "").replace("bbb_", "")) - Integer.parseInt(o2.replace("aaa_", "").replace("bbb_", ""));
//            }
//        });
//		Collections.sort(l);
    for (String names : l) {
      if (names.indexOf("bbb") >= 0) {
        // 添加标题
        String _ids = request.getParameter(names).toString();
        Map<String, Object> map1 = new HashMap<String, Object>();
        if ("20".equals(_ids)) {
          map1.put("content",
              list2.get(18).get("CONNT") == null ? "" : list2.get(18).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省时户数消耗情况");

          XWPFTable table = WorderToNewWordUtils.setTable(map1, document, new String[][] {
                  new String[] { "序号", "单位", "等效总;用户数", "累计用户平均停电情况", "", "", "", "停电时户数消耗情况", "", "", "" },
                  new String[] { "", "", "", "全口;径", "同比", "预;安排;停电", "故障;停电", "年度;目标", "累计;消耗", "同比",
                      "本周;消耗" } },
              getREPORT_TDSHS(obj));

          WorderToNewWordUtils.mergeCellsHorizontal(table, 0, 3, 6);
          WorderToNewWordUtils.mergeCellsHorizontal(table, 0, 7, 10);
          WorderToNewWordUtils.mergeCellsHorizontal(table, 2, 0, 1);

          WorderToNewWordUtils.mergeCellsVertically(table, 0, 0, 1);
          WorderToNewWordUtils.mergeCellsVertically(table, 1, 0, 1);
          WorderToNewWordUtils.mergeCellsVertically(table, 2, 0, 1);
        } else if ("21".equals(_ids)) {

          map1.put("content", "");
          name_index++;
          map1.put("tabletilte", "全省停电异常情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document, new String[][] {
                  new String[] { "序号", "单位", "大范围停电情况", "", "", "重复停电情况", "", "", "", "单户超长时间情况", "" },
                  new String[] { "", "", "本周发生", "累计发生", "占比", "累计3-4;次", "占比", "累计5次;及以上", "占比", "用户数(>24)",
                      "最多停电;	次数" } },
              getREPORT_TDYCQK(obj));

          WorderToNewWordUtils.mergeCellsHorizontal(table2, 0, 2, 4);
          WorderToNewWordUtils.mergeCellsHorizontal(table2, 0, 5, 8);
          WorderToNewWordUtils.mergeCellsHorizontal(table2, 0, 9, 10);
          WorderToNewWordUtils.mergeCellsHorizontal(table2, 2, 0, 1);

          WorderToNewWordUtils.mergeCellsVertically(table2, 0, 0, 1);
          WorderToNewWordUtils.mergeCellsVertically(table2, 1, 0, 1);
        } else if ("22".equals(_ids)) {

          map1.put("content",
              list2.get(19).get("CONNT") == null ? "" : list2.get(19).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省配网故障情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] {
                  new String[] { "序号", "单位", "重合不良", "重合良好", "接地障碍", "分支线故障", "合计", "同比", "环比" } },
              getREPORT_PWGZQK(obj));

          WorderToNewWordUtils.mergeCellsHorizontal(table2, 1, 0, 1);
        } else if ("25".equals(_ids)) {

          map1.put("content",
              list2.get(20).get("CONNT") == null ? "" : list2.get(20).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省低电压、三相不平衡情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document, new String[][] {

                  new String[] { "序号", "单位", "公变数(台)", "低电压", "", "", "三相不平衡", "", "" },
                  new String[] { "", "", "", "累计", "本周", "时长(小时)", "累计", "本周", "时长(小时)" } },

              getREPORT_DDYSXBPH(obj));
          WorderToNewWordUtils.mergeCellsVertically(table2, 0, 0, 1);
          WorderToNewWordUtils.mergeCellsVertically(table2, 1, 0, 1);
          WorderToNewWordUtils.mergeCellsVertically(table2, 2, 0, 1);

          WorderToNewWordUtils.mergeCellsHorizontal(table2, 0, 3, 5);
          WorderToNewWordUtils.mergeCellsHorizontal(table2, 0, 6, 8);
          WorderToNewWordUtils.mergeCellsHorizontal(table2, 2, 0, 1);
        } else if ("26".equals(_ids)) {

          map1.put("content", "");
          name_index++;
          map1.put("tabletilte", "全省重过载情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] { new String[] { "序号", "单位", "公变数(台)", "重过载", "", "", "", "", "" },
                  new String[] { "", "", "", "累计", "过载台数", "时长(小时)", "累计", "重载台数", "时长(小时)" } },

              getREPORT_ZGZQK(obj));
          WorderToNewWordUtils.mergeCellsVertically(table2, 0, 0, 1);
          WorderToNewWordUtils.mergeCellsVertically(table2, 1, 0, 1);
          WorderToNewWordUtils.mergeCellsVertically(table2, 2, 0, 1);

          WorderToNewWordUtils.mergeCellsHorizontal(table2, 0, 3, 8);
          WorderToNewWordUtils.mergeCellsHorizontal(table2, 2, 0, 1);
        } else if ("27".equals(_ids)) {

          map1.put("content",
              list2.get(21).get("CONNT") == null ? "" : list2.get(21).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省配网报修情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] {
                  new String[] { "序号", "单位", "累计报修数量", "同比", "本周报修数量", "", "", "", "", "", "" },
                  new String[] { "", "", "", "", "高压故障", "低压故障", "计量故障", "客户内部故障", "电能质量故障", "非电力故障",
                      "合计" } },
              getREPORT_PWBXQK(obj));
          WorderToNewWordUtils.mergeCellsVertically(table2, 0, 0, 1);
          WorderToNewWordUtils.mergeCellsVertically(table2, 1, 0, 1);
          WorderToNewWordUtils.mergeCellsVertically(table2, 2, 0, 1);
          WorderToNewWordUtils.mergeCellsVertically(table2, 3, 0, 1);

          WorderToNewWordUtils.mergeCellsHorizontal(table2, 0, 4, 10);
        } else if ("28".equals(_ids)) {

          map1.put("content",
              list2.get(22).get("CONNT") == null ? "" : list2.get(22).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省配网报修投诉情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document, new String[][] {
                  new String[] { "序号", "单位", "累计生产类投诉", "本周生产类重点关注投诉", "", "", "", "", "", "", "", "" },
                  new String[] { "", "", "", "抢修超时限", "抢修人员服务态度", "抢修人员服务规范", "抢修质量", "频繁停电", "电压质量长时间异常",
                      "供电频率长时间异常", "合计", "同比" } },

              getREPORT_PWBXTSQK(obj));

          WorderToNewWordUtils.mergeCellsVertically(table2, 0, 0, 1);
          WorderToNewWordUtils.mergeCellsVertically(table2, 1, 0, 1);
          WorderToNewWordUtils.mergeCellsVertically(table2, 2, 0, 1);

          WorderToNewWordUtils.mergeCellsHorizontal(table2, 0, 3, 11);
          WorderToNewWordUtils.mergeCellsHorizontal(table2, 2, 0, 1);
        } else if ("29".equals(_ids)) {

          map1.put("content",
              list2.get(23).get("CONNT") == null ? "" : list2.get(23).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省高负损情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] { new String[] { "序号", "单位", "线路条数", "负损、高损线路情况", "", "" },
                  new String[] { "", "", "", "负损线路条数", "高损线路条数【10-30】", "超大损线路条数【>30】" } },

              getREPORT_GFSQK(obj));

          WorderToNewWordUtils.mergeCellsVertically(table2, 0, 0, 1);
          WorderToNewWordUtils.mergeCellsVertically(table2, 1, 0, 1);
          WorderToNewWordUtils.mergeCellsVertically(table2, 2, 0, 1);

          WorderToNewWordUtils.mergeCellsHorizontal(table2, 0, 3, 5);
        } else if ("1".equals(_ids)) {
          map1.put("content", list2.get(0).get("CONNT") == null ? "" : list2.get(0).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省作业计划执行情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] {
                  new String[] { "单位", "10千伏停电计划（项）", "实际执行计划", "执行率", "延时停、送电的计划", "临时计划" } },

              getYYFX_PWTD(obj), new String[] {
                  "注：统计周期为报表当期。计划未执行原因包括：1.天气影响，2.施工受阻，3.作业准备不足，4.客户原因，5.上级电网影响，6.其他原因请具体说明；计划停、送电延时原因包括：1.天气影响，2.施工受阻，3.作业准备不足，4.客户原因，5.上级电网影响，6.其他原因请具体说明。" });

        } else if ("2".equals(_ids)) {
          map1.put("content", list2.get(1).get("CONNT") == null ? "" : list2.get(1).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省频繁停电线路数量");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] { new String[] { "单位", "频繁停电线路;（条）", "近两月停电5 次及以上;线路（条）", "近两月重复计划停电;线路（条）",
                  "近两月故障停电两次及;以上线路（条）", "计划作业后两个月内再次发生故障;停电线路（条）" } },
              getYYFX_PFTD(obj), new String[] {
                  "注：1.统计周期为报表当期；2.统计周期为近两月内发生3 次及以上停电；3.频繁停电判定：统计包含计划停电、故障跳闸重合不良、接地故障处置（不含选线过程），紧急停电（含消缺）等全部停电事件。" });

        } else if ("3".equals(_ids)) {
          map1.put("content", list2.get(2).get("CONNT") == null ? "" : list2.get(2).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省重过载线路数量");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] { new String[] { "单位", "重过载线路（条）", "重载线路", "过载线路", "治理完成" } },
              getYYFX_ZGZXL(obj), new String[] {
                  "注：1.统计周期为报表当期；2.重载线路判定：负载率达到70%-100%且持续1 小时及以上的线路；3.过载线路判定：负载率大于100%且持续1 小时及以上的线路。" });

        } else if ("4".equals(_ids)) {
          map1.put("content", list2.get(3).get("CONNT") == null ? "" : list2.get(3).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省配变异常数量");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] { new String[] { "单位", "异常配变数量（台）", "重载", "过载", "低电压", "三相不平衡" } },
              getYYFX_YCSB(obj), new String[] { "注：统计周期为报表当期。数据取自于供电服务指挥系统。" });

        } else if ("5".equals(_ids)) {
          map1.put("content", list2.get(4).get("CONNT") == null ? "" : list2.get(4).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "配变异常监测预警情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document, new String[][] {
                  new String[] { "单位", "配变异常预警工单（件）", "已归档（件）", "治理完成异常配变（台）", "建议纳入工程治理（台）", "督办工单（张）" } },
              getYYFX_PBYC(obj), new String[] { "注：统计周期为报表当期。" });
        } else if ("6".equals(_ids)) {
          map1.put("content", list2.get(5).get("CONNT") == null ? "" : list2.get(5).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省巡视、消缺计划执行情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] { new String[] { "单位", "巡视计划数量", "已执行", "发现缺陷数", "", "", "消除缺陷数", "", "" },

                  new String[] { "", "", "", "危急缺陷", "危急缺陷", "一般缺陷", "危急缺陷", "危急缺陷", "一般缺陷" },

              }, getYYFX_XQJH(obj), new String[] { "注：统计周期为报表当期。数据取自于供电服务指挥系统。" });
          WorderToNewWordUtils.mergeCellsVertically(table2, 0, 0, 1);
          WorderToNewWordUtils.mergeCellsVertically(table2, 1, 0, 1);
          WorderToNewWordUtils.mergeCellsVertically(table2, 2, 0, 1);

          WorderToNewWordUtils.mergeCellsHorizontal(table2, 0, 3, 5);
          WorderToNewWordUtils.mergeCellsHorizontal(table2, 0, 6, 8);
        } else if ("7".equals(_ids)) {
          map1.put("content", list2.get(6).get("CONNT") == null ? "" : list2.get(6).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省操作票执行情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] { new String[] { "单位", "编制操作票（份）", "执行操作票（份）", "执行率" }, },
              getYYFX_CZPZX(obj),
              new String[] { "注：统计周期为报表当期。未执行原因包括：1.天气影响，2.施工受阻，3.作业准备不足，4.客户原因，5.上级电网影响，6.其他原因请具体说明；" });
        } else if ("8".equals(_ids)) {
          map1.put("content", list2.get(7).get("CONNT") == null ? "" : list2.get(7).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省新投运设备数量");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] { new String[] { "单位", "10 千伏配电线", "", "公用配电变压器", "" },
                  new String[] { "单位", "数量（条）", "长度（公里）", "数量（台）", "容量（MVA）" }, },
              getYYFX_XSB(obj), new String[] { "注：统计周期为报表当期。" });
          WorderToNewWordUtils.mergeCellsVertically(table2, 0, 0, 1);

          WorderToNewWordUtils.mergeCellsHorizontal(table2, 0, 1, 2);
          WorderToNewWordUtils.mergeCellsHorizontal(table2, 0, 3, 4);
        } else if ("9".equals(_ids)) {
          map1.put("content", list2.get(8).get("CONNT") == null ? "" : list2.get(8).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省故障抢修服务情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] {
                  new String[] { "单位", "抢修工单数（件）", "移动APP接单率", "抢修平均到达现场时长（分钟）", "故障平均修复时长（分钟）" }, },
              getYYFX_QXQK(obj), new String[] { "注：统计周期为报表当期。数据取自于供电服务指挥系统。" });
        } else if ("10".equals(_ids)) {
          map1.put("content", list2.get(9).get("CONNT") == null ? "" : list2.get(9).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省停电信息报送情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] { new String[] { "单位", "发布停电信息（条）", "停电信息发布及时率", "停电范围分析到户率" }, },
              getYYFX_TDBS(obj), new String[] { "注：统计周期为报表当期。数据取自于供电服务指挥系统。" });
        } else if ("11".equals(_ids)) {
          map1.put("content",
              list2.get(10).get("CONNT") == null ? "" : list2.get(10).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "95598 服务指标情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] { new String[] { "单位", "营销类考核投诉件数", "生产类考核投诉件数", "95598 工单处理不及时件数",
                  "95598 业务处理不规范件数	", "客户三次致电件数" }, },
              getYYFX_FWZB(obj),
              new String[] { "注：表中各项指标统计周期为报表当期。数据取自于供电服务指挥系统。", "营销类及生产类考核投诉件数指同业对标中客户服务满意率中考核范围内的投诉件数。",
                  "（1）营销服务不规范投诉包括20 项三级投诉：1.营业厅服务，2.营业厅人员服务态度，3.营业厅人员服务规范，4.收费标准，5.收费项目，6.业扩报装超时限，7.环节处理不当，8.业务办理超时限，9.环节处理问题，10.抄表，11.勘测人员服务态度，12.勘测人员服务规范，13.电价，14.电费，15.计量人员服务态度，16.计量人员服务规范，17.表计线路接错，18.计量装置，19.户表轮换改造，20.验表。",
                  "（2）非营销服务不规范投诉包括7 项三级投诉：1.抢修超时限，2.抢修人员服务态度，3.抢修人员服务规范，4.抢修质量，5.频繁停电，6.电压质量长时间异常，7.供电频率长时间异常。" });
        } else if ("12".equals(_ids)) {
          map1.put("content",
              list2.get(11).get("CONNT") == null ? "" : list2.get(11).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省营业厅监控情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] { new String[] { "单位", "营业厅数量（个）	", "可监控营业厅", "本期巡查营业厅", "发现不规范行为（处）" }, },
              getYYFX_JKQK(obj), new String[] { "注：统计周期为报表当期。" });
        } else if ("13".equals(_ids)) {
          map1.put("content",
              list2.get(12).get("CONNT") == null ? "" : list2.get(12).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "10 千伏业扩全流程时长管控情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] { new String[] { "单位", "超过70 天少于120 天未送电数", "超过120 天少于300 天未送电数",
                  "超过300 天未送电数", "超过70 天未送电合计数" }, },
              getYYFX_YKQLC(obj), new String[] { "注：统计周期为报表当期。" });
        } else if ("14".equals(_ids)) {
          map1.put("content",
              list2.get(13).get("CONNT") == null ? "" : list2.get(13).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "业扩报装全流程管控情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document, new String[][] {
                  new String[] { "单位", "10 千伏项目平均接电时间", "400 伏项目平均接电时间", "主环节超期件数", "督办预警次数", "督办预警消除次数" }, },
              getYYFX_YKBZ(obj), new String[] { "注：10 千伏、400 伏项目平均接电时间，统计周期为报表当期。" });
        } else if ("15".equals(_ids)) {
          map1.put("content",
              list2.get(14).get("CONNT") == null ? "" : list2.get(14).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "全省低压业扩报装接入容量情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document, new String[][] {
                  new String[] { "单位", "城市20-100（160）千瓦新装（件、%）", "", "", "农村20-80（100）千瓦新装（件、%）", "", "" },
                  new String[] { "单位", "受理工单数", "低压接入数", "低压接入率", "受理工单数", "低压接入数", "低压接入率" }, },
              getYYFX_DYYK(obj), new String[] { "注：统计周期为报表当期。" });

          WorderToNewWordUtils.mergeCellsHorizontal(table2, 0, 1, 3);
          WorderToNewWordUtils.mergeCellsHorizontal(table2, 0, 4, 6);

          WorderToNewWordUtils.mergeCellsVertically(table2, 0, 0, 1);
        } else if ("16".equals(_ids)) {
          map1.put("content",
              list2.get(15).get("CONNT") == null ? "" : list2.get(15).get("CONNT").toString());
          name_index++;
          map1.put("tabletilte", "营配调数据治理情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document, new String[][] {
                  new String[] { "单位", "核查营配调问题数据（条）", "下发问题工单（件）", "处理问题数据（条）", "处理率", "现场核查线路（条）" }, },
              getYYFX_YPD(obj), new String[] { "注：统计周期为报表当期。" });
        } else if ("19".equals(_ids)) {
          map1.put("content", "");
          name_index++;
          map1.put("tabletilte", "新投运设备明细（城网）");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] { new String[] { "序号", "单位", "设备名称", "设备类型", "PMS系统设备编码", "投运日期" },

              }, getYYFX_XTY(obj));
        } else if ("18".equals(_ids)) {
          map1.put("content", "");
          name_index++;
          map1.put("tabletilte", "全省重过载线路明细");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] {
                  new String[] { "序号", "单位", "线路名称", "额定电流", "最大电流", "最大负载率", "重/过载持续时长" }, },
              getYYFX_ZGZMX(obj));
        } else if ("17".equals(_ids)) {
          map1.put("content", "");
          name_index++;
          map1.put("tabletilte", "全省频繁停电线路明细");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document, new String[][] { new String[] {
                  "序号", "单位", "线路名称", "近两月停电次数", "其中近两月故障停电次数", "其中近两月计划停电次数", "其中近两月临时停电次数" }, },
              getYYFX_PFTDMX(obj));
        } else if ("23".equals(_ids)) {
          map1.put("content", "");
          name_index++;
          map1.put("tabletilte", "本周两次及以上故障重复停电情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] { new String[] { "序号", "单位", "线路名称", "停电原因", "停电时长" }, },
              getREPORT_LCJYSGZTD(obj));
        } else if ("24".equals(_ids)) {
          map1.put("content", "");
          name_index++;
          map1.put("tabletilte", "本周8小时以上故障停电情况");
          XWPFTable table2 = WorderToNewWordUtils.setTable(map1, document,
              new String[][] { new String[] { "序号", "单位", "线路名称", "停电原因", "停电时长" }, },
              getREPORT_BXSYSGZTD(obj));
        }
      }

      if (names.indexOf("aaa") >= 0) {
        // 添加表格
        WorderToNewWordUtils.setTitle(document, request.getParameter(names).toString());
      }
    }

    int i = 0;
//			if("1".equals(_ids)){
//				Map<String, Object> map1 = new HashMap<String, Object>();
//				map1.put("title", name[name_index] + "天啊");
//				map1.put("tabletilte", "表格标题");
//				map1.put("content", list2.get(0).get("CONNT") == null ? "" : list2.get(0).get("CONNT").toString());
//
//				WorderToNewWordUtils.setTable(map1, document,
//
//						new String[][] { new String[] { "单位", "10千伏停;电计划（项）", "实际执行;计划", "执行率", "延时停、送电;的计划", "临时计划" } }
//
//						, getFirst(obj));
//				name_index++;
//			}

    // REPORT_TDSHS 全省时户数消耗情况时户数

    try {

      ServletOutputStream ServletOutputStream = response.getOutputStream();
      response.setContentType("application/octet-stream");
      response.setHeader("Content-Disposition",
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
 */
