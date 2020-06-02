package com.xx.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

/**
 * @author aqi
 * DateTime: 2020/5/20 11:03 上午
 * Description: No Description
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class User {

    private Integer id;
    private String name;
    private String password;
    private Integer gender;
    private Boolean live;
    private String remarks;
    private Date createTime;
    private String other;
    private String msg1;
    private String msg2;
    private String msg3;

}
