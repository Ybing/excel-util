package com.bing.studyexcel.util;

import java.lang.annotation.*;

/**
 * @Description: TODO
 * @Author: 杨亚兵
 * @Date: 2019/10/30 14:12
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Excel {

    /**
     * 对应的列名
     */
    String value() default "";
    /**
     * 顺序
     * */
    int order() default 0;
    /**
     * 导入用  是否必填
     * */
    boolean required() default false;
}
