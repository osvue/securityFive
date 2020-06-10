package org.javaboy.formlogin.config.security;

import com.fasterxml.jackson.databind.ObjectMapper;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.HashMap;
import javax.servlet.FilterChain;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.springframework.security.core.Authentication;
import org.springframework.security.web.authentication.AuthenticationSuccessHandler;
import org.springframework.stereotype.Component;

/**
 * @Author: Mr.Han
 * @Description: securityFive
 * @Date: Created in 2020/6/10_10:57
 * @Modified By: THE GIFTED
 */
@Component
public class SuccessLogin implements AuthenticationSuccessHandler {

  @Override
  public void onAuthenticationSuccess(HttpServletRequest request, HttpServletResponse response,
      FilterChain chain, Authentication authentication) throws IOException, ServletException {

  }

  @Override
  public void onAuthenticationSuccess(HttpServletRequest request, HttpServletResponse resp,
      Authentication authentication) throws IOException, ServletException {
    Object principal = authentication.getPrincipal();
    resp.setContentType("application/json;charset=utf-8");
    PrintWriter out = resp.getWriter();
    HashMap<String,Object> map = new HashMap<>();
    map.put("data",principal);
    map.put("code",20000);

    out.write(new ObjectMapper().writeValueAsString(map));
    out.flush();
    out.close();
  }
}
 