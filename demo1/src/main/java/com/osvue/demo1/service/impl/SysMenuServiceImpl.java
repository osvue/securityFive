package com.osvue.demo1.service.impl;

import org.springframework.stereotype.Service;
import javax.annotation.Resource;
import java.util.List;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import com.osvue.demo1.domain.SysMenu;
import java.util.List;
import com.osvue.demo1.domain.SysMenuExample;
import com.osvue.demo1.mapper.SysMenuMapper;
import com.osvue.demo1.service.SysMenuService;
/**
 * @Author: Mr.Han
 * @Description:  securityFive
 * @Date: Created in 2020/6/11_13:52
 * @Modified By: THE GIFTED 
 */
@Service
public class SysMenuServiceImpl extends ServiceImpl<SysMenuMapper, SysMenu> implements SysMenuService{

    @Override
    public long countByExample(SysMenuExample example) {
        return baseMapper.countByExample(example);
    }
    @Override
    public int deleteByExample(SysMenuExample example) {
        return baseMapper.deleteByExample(example);
    }
    @Override
    public List<SysMenu> selectByExample(SysMenuExample example) {
        return baseMapper.selectByExample(example);
    }
    @Override
    public int updateByExampleSelective(SysMenu record,SysMenuExample example) {
        return baseMapper.updateByExampleSelective(record,example);
    }
    @Override
    public int updateByExample(SysMenu record,SysMenuExample example) {
        return baseMapper.updateByExample(record,example);
    }
}
