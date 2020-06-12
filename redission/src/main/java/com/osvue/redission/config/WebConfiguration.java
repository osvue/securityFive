package com.osvue.redission.config;

/**
 * @Author: Mr.Han @Description: securityFive @Date: Created in 2020/6/12_14:48 @Modified By: THE
 * GIFTED
 */
import com.osvue.redission.interceptor.AutoIdempotentInterceptor;
import javax.annotation.Resource;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.InterceptorRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

@Configuration
public class WebConfiguration implements WebMvcConfigurer {

  @Resource
  private AutoIdempotentInterceptor autoIdempotentInterceptor;

  /**
   * 添加拦截器
   *
   * @param registry
   */
  @Override
  public void addInterceptors(InterceptorRegistry registry) {
    registry.addInterceptor(autoIdempotentInterceptor).addPathPatterns("/**")
    .excludePathPatterns("/js/**","/css/**","/image/**");

  }
}
