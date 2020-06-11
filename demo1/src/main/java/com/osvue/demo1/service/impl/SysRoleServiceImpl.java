package com.osvue.demo1.service.impl;

import org.springframework.stereotype.Service;
import javax.annotation.Resource;
import java.util.List;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import java.util.List;
import com.osvue.demo1.mapper.SysRoleMapper;
import com.osvue.demo1.domain.SysRoleExample;
import com.osvue.demo1.domain.SysRole;
import com.osvue.demo1.service.SysRoleService;
/**
 * @Author: Mr.Han
 * @Description:  securityFive
 * @Date: Created in 2020/6/11_13:53
 * @Modified By: THE GIFTED 
 */
@Service
public class SysRoleServiceImpl extends ServiceImpl<SysRoleMapper, SysRole> implements SysRoleService{

    @Override
    public long countByExample(SysRoleExample example) {
        return baseMapper.countByExample(example);
    }
    @Override
    public int deleteByExample(SysRoleExample example) {
        return baseMapper.deleteByExample(example);
    }
    @Override
    public List<SysRole> selectByExample(SysRoleExample example) {
        return baseMapper.selectByExample(example);
    }
    @Override
    public int updateByExampleSelective(SysRole record,SysRoleExample example) {
        return baseMapper.updateByExampleSelective(record,example);
    }
    @Override
    public int updateByExample(SysRole record,SysRoleExample example) {
        return baseMapper.updateByExample(record,example);
    }
}
