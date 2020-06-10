package com.osvue.demo1;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * @Author: Mr.Han
 * @Description: securityFive
 * @Date: Created in 2020/5/23_10:47
 * @Modified By: THE GIFTED
 */
@RestController
public class DemoController {

  @GetMapping(value = "/say")
  public String say(){
    return "Hello World";
  }
}
 