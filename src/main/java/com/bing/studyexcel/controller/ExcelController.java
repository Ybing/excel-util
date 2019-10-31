package com.bing.studyexcel.controller;

import com.bing.studyexcel.pojo.User;
import com.bing.studyexcel.util.ExcelUtil;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

/**
 * @Description: TODO
 * @Author: 杨亚兵
 * @Date: 2019/10/30 14:55
 */
@Controller
public class ExcelController {

    @ResponseBody
    @RequestMapping("/import_excel")
    public String importExcel(MultipartFile file) throws Exception {
        List<User> userList = ExcelUtil.importExcel(file, User.class);
        return userList.toString();
    }

    @ResponseBody
    @RequestMapping("/export_excel")
    public void exportExcel(HttpServletResponse response) throws Exception {

        List<User>userList = new LinkedList<>();
        User user1 = new User();
        User user2 = new User();
        user1.setUserName("小王");
        user2.setUserName("tom");
        user1.setAge(22);
        user2.setAge(22);
        user1.setBirthday(new Date());
        user2.setBirthday(new Date());
        user1.setMoney(12345D);
        user2.setMoney(67890D);
        userList.add(user1);
        userList.add(user2);

        String baseName = "员工表";
        String extension = ExcelUtil.EXCEL_XLS;

        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        response.addHeader("Content-Disposition", "attachment; filename=" + new String(baseName.getBytes("gb2312"), "iso8859-1") + "." + extension);
        ExcelUtil.exportExcel(ExcelUtil.EXCEL_XLS,"员工表",null,userList,User.class,response.getOutputStream());
    }
}
