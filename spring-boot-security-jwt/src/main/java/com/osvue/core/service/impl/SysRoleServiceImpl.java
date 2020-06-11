package com.osvue.core.service.impl;

import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import com.osvue.core.dao.SysRoleDao;
import com.osvue.core.entity.SysRoleEntity;
import com.osvue.core.service.SysRoleService;
import org.springframework.stereotype.Service;

/**
 * @Description 角色业务实现
 * @Author Osvue.gitee.io
 * @CreateTime 2019/9/14 15:57
 */
@Service("sysRoleService")
public class SysRoleServiceImpl extends ServiceImpl<SysRoleDao, SysRoleEntity> implements SysRoleService {

}