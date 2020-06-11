package com.osvue.demo1.service;

import com.osvue.demo1.domain.SysMenu;
import java.util.List;
import com.osvue.demo1.domain.SysMenuExample;
import com.baomidou.mybatisplus.extension.service.IService;
    /**
 * @Author: Mr.Han
 * @Description:  securityFive
 * @Date: Created in 2020/6/11_13:52
 * @Modified By: THE GIFTED 
 */
public interface SysMenuService extends IService<SysMenu>{


    long countByExample(SysMenuExample example);

    int deleteByExample(SysMenuExample example);

    List<SysMenu> selectByExample(SysMenuExample example);

    int updateByExampleSelective(SysMenu record,SysMenuExample example);

    int updateByExample(SysMenu record,SysMenuExample example);

}
