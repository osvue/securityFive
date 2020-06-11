package com.osvue.core.entity;

import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import lombok.Data;
import java.io.Serializable;

/**
 * @Description 权限实体
 * @Author Osvue.gitee.io
 * @CreateTime 2019/9/14 15:57
 */
@Data
@TableName("sys_menu")
public class SysMenuEntity implements Serializable {
	private static final long serialVersionUID = 1L;

	/**
	 * 权限ID
	 */
	@TableId
	private Long menuId;
	/**
	 * 权限名称
	 */
	private String name;
	/**
	 * 权限标识
	 */
	private String permission;

}
