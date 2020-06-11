package com.osvue.demo1.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import com.osvue.demo1.domain.SysMenu;
import com.osvue.demo1.domain.SysMenuExample;
import java.util.List;
import org.apache.ibatis.annotations.Param;

/**
 * @Author: Mr.Han
 * @Description:  securityFive
 * @Date: Created in 2020/6/11_13:52
 * @Modified By: THE GIFTED 
 */
public interface SysMenuMapper extends BaseMapper<SysMenu> {
    long countByExample(SysMenuExample example);

    int deleteByExample(SysMenuExample example);

    List<SysMenu> selectByExample(SysMenuExample example);

    int updateByExampleSelective(@Param("record") SysMenu record, @Param("example") SysMenuExample example);

    int updateByExample(@Param("record") SysMenu record, @Param("example") SysMenuExample example);
}