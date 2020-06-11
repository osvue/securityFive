package com.osvue.demo1.domain;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import java.io.Serializable;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @Author: Mr.Han
 * @Description:  securityFive
 * @Date: Created in 2020/6/11_13:52
 * @Modified By: THE GIFTED 
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@TableName(value = "sys_menu")
public class SysMenu implements Serializable {
    /**
     * ID
     */
    @TableId(value = "menu_id", type = IdType.AUTO)
    private Long menuId;

    /**
     * 权限名称
     */
    @TableField(value = "name")
    private String name;

    /**
     * 权限标识
     */
    @TableField(value = "permission")
    private String permission;

    private static final long serialVersionUID = 1L;

    public static final String COL_MENU_ID = "menu_id";

    public static final String COL_NAME = "name";

    public static final String COL_PERMISSION = "permission";
}