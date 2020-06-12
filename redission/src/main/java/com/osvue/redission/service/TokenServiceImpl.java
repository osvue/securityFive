package com.osvue.redission.service;

import com.osvue.redission.utils.Constant;
import com.osvue.redission.utils.RedisUtil;
import java.util.UUID;
import javax.servlet.http.HttpServletRequest;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

/**
 * @Author: Mr.Han
 * @Description: securityFive
 * @Date: Created in 2020/6/12_14:01
 * @Modified By: THE GIFTED
 */
@Service
public class TokenServiceImpl implements TokenService{

  @Autowired
  private RedisUtil redisService;


  /**
   * 创建token
   *
   * @return
   */
  @Override
  public String createToken() {
    String str = UUID.randomUUID().toString();
    StringBuilder token = new StringBuilder();
    try {
      token.append("RT_").append(str);
      redisService.setEx(token.toString(), token.toString(),10000L);
      boolean notEmpty = StringUtils.isNotEmpty(token.toString());
      if (notEmpty) {
        return token.toString();
      }
    }catch (Exception ex){
      ex.printStackTrace();
    }
    return null;
  }


  /**
   * 检验token
   *
   * @param request
   * @return
   */
  @Override
  public boolean checkToken(HttpServletRequest request) throws Exception {

    String token = request.getHeader(Constant.TOKEN_NAME);
    if (StringUtils.isBlank(token)) {// header中不存在token
      token = request.getParameter(Constant.TOKEN_NAME);
      if (StringUtils.isBlank(token)) {// parameter中也不存在token
        throw new RuntimeException ("ILLEGAL_ARGUMENT, 100");
      }
    }

    if (!redisService.exists(token)) {
      throw new RuntimeException("ILLEGAL_ARGUMENT, 100");
    }

    boolean remove = redisService.remove(token);
    if (!remove) {
      throw new RuntimeException(  "ILLEGAL_ARGUMENT, 100");
    }
    return true;
  }
}
 