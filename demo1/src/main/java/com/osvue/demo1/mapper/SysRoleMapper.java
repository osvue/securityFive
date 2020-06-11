package com.osvue.demo1.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import com.osvue.demo1.domain.SysRole;
import com.osvue.demo1.domain.SysRoleExample;
import java.util.List;
import org.apache.ibatis.annotations.Param;

/**
 * @Author: Mr.Han
 * @Description:  securityFive
 * @Date: Created in 2020/6/11_13:53
 * @Modified By: THE GIFTED 
 */
public interface SysRoleMapper extends BaseMapper<SysRole> {
    long countByExample(SysRoleExample example);

    int deleteByExample(SysRoleExample example);

    List<SysRole> selectByExample(SysRoleExample example);

    int updateByExampleSelective(@Param("record") SysRole record, @Param("example") SysRoleExample example);

    int updateByExample(@Param("record") SysRole record, @Param("example") SysRoleExample example);
}