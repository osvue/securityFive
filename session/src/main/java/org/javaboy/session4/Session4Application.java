package org.javaboy.session4;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
//// 这里的./conf/nginx.conf ，最前面的.代表当前目录，因为我是在nginx安装目录执行的，所以最前面的这个.不能缺少，否则找不到nginx.conf
// nginx -c ./conf/nginx.conf
@SpringBootApplication
public class Session4Application {

    public static void main(String[] args) {
        SpringApplication.run(Session4Application.class, args);
    }

}
