package org.osvue.formlogin;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * @Author: Mr.Han
 * @Description: securityFive
 * @Date: Created in 2020/5/23_10:45
 * @Modified By: THE GIFTED
 */
@RestController
public class HelloController {
    @GetMapping("/hello")
    public String hello() {
        return "hello";
    }
}
