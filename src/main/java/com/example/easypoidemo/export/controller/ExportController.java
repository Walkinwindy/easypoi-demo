package com.example.easypoidemo.export.controller;

import com.example.easypoidemo.export.test.MyExportTest;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

/**
 * 一句话功能描述
 *
 * @author zhaoXinXing
 * @date 2019/10/29 14:37
 */
@Controller
public class ExportController {
    @Autowired
    MyExportTest myExportTest;

    /**
     * 普通导出 学生成绩表
     *
     * @param null
     * @author zhaoXinXing
     * @date 2019/10/29
     */
    @GetMapping("export1")
    public void getExcel1() {
        myExportTest.test1111();
    }
    /**
     * 模板导出
     * @author zhaoXinXing
     * @date 2019/10/29
     * @param null
     * @return
     */
    @GetMapping("export2")
    public void getExcel2() {
        myExportTest.test1112();
    }

    @RequestMapping("/")
    public String getIndex(){
        return "index";
    }

}
