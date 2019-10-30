package com.bing.studyexcel.pojo;

import com.bing.studyexcel.util.Excel;
import lombok.Data;

import java.util.Date;

/**
 * @Description: TODO
 * @Author: 杨亚兵
 * @Date: 2019/10/30 14:57
 */
@Data
public class User {
    @Excel(value = "姓名",order = 1)
    private String userName;
    @Excel(value = "年龄",order = 2)
    private Integer age;
    @Excel(value = "生日",order = 3)
    private Date birthday;
}
