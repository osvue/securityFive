package com.osvue.redission.controller;

import com.osvue.redission.anno.AutoIdempotent;
import com.osvue.redission.service.TokenService;
import com.osvue.redission.utils.ResultBean;
import javax.annotation.Resource;
import org.apache.commons.lang3.StringUtils;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * @Author: Mr.Han
 * @Description: securityFive
 * @Date: Created in 2020/6/12_14:47
 * @Modified By: THE GIFTED
 */
@RestController
public class BusinessController {
  @Resource
  private TokenService tokenService;



  @PostMapping("/get/token")
  public ResultBean<String> getToken(){
    String token = tokenService.createToken();
    if (StringUtils.isNotEmpty(token)) {

      return new ResultBean<>(token);
    }
    return new ResultBean<>("StrUtil.EMPTY");
  }


  @AutoIdempotent
  @PostMapping("/test")
  public String testIdempotence() {


    return "SUCCESS";
  }
}
 