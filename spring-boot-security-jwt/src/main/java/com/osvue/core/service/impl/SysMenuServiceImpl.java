package com.osvue.core.service.impl;

import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import com.osvue.core.dao.SysMenuDao;
import com.osvue.core.entity.SysMenuEntity;
import com.osvue.core.service.SysMenuService;
import org.springframework.stereotype.Service;

/**
 * @Description 权限业务实现
 * @Author Osvue.gitee.io
 * @CreateTime 2019/9/14 15:57
 */
@Service("sysMenuService")
public class SysMenuServiceImpl extends ServiceImpl<SysMenuDao, SysMenuEntity> implements SysMenuService {

}