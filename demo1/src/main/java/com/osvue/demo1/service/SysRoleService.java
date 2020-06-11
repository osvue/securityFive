package com.osvue.demo1.service;

import java.util.List;
import com.osvue.demo1.domain.SysRoleExample;
import com.osvue.demo1.domain.SysRole;
import com.baomidou.mybatisplus.extension.service.IService;
    /**
 * @Author: Mr.Han
 * @Description:  securityFive
 * @Date: Created in 2020/6/11_13:53
 * @Modified By: THE GIFTED 
 */
public interface SysRoleService extends IService<SysRole>{


    long countByExample(SysRoleExample example);

    int deleteByExample(SysRoleExample example);

    List<SysRole> selectByExample(SysRoleExample example);

    int updateByExampleSelective(SysRole record,SysRoleExample example);

    int updateByExample(SysRole record,SysRoleExample example);

}
