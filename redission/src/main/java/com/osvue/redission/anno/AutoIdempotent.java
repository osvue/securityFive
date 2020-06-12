package com.osvue.redission.anno;

/**
 * @Author: Mr.Han
 * @Description: securityFive
 * @Date: Created in 2020/6/12_13:49
 * @Modified By: THE GIFTED
 */
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 在需要保证 接口幂等性 的Controller的方法上使用此注解
 */
@Target({ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
public @interface AutoIdempotent {

}