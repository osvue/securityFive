package com.osvue.demo1.utils;

import com.alibaba.fastjson.JSON;
import com.osvue.demo1.config.JWTConfig;
import com.osvue.demo1.domain.SysUser;
import io.jsonwebtoken.Jwts;
import io.jsonwebtoken.SignatureAlgorithm;
import java.util.Date;

/**
 * @Author: Mr.Han
 * @Description: securityFive
 * @Date: Created in 2020/6/11_13:54
 * @Modified By: THE GIFTED
 */

public class JWTTokenUtil {
  public static String createAccessToken(SysUser selfUserEntity){
    // 登陆成功生成JWT
    String token = Jwts.builder()
        // 放入用户名和用户ID
        .setId(selfUserEntity.getUserId()+"")
        // 主题
        .setSubject(selfUserEntity.getUsername())
        // 签发时间
        .setIssuedAt(new Date())
        // 签发者
        .setIssuer("sans")
        // 自定义属性 放入用户拥有权限
        .claim("authorities", JSON.toJSONString(null))
        // 失效时间
        .setExpiration(new Date(System.currentTimeMillis() + JWTConfig.expiration))
        // 签名算法和密钥
        .signWith(SignatureAlgorithm.HS512, JWTConfig.secret)
        .compact();
    return token;
  }
}
 