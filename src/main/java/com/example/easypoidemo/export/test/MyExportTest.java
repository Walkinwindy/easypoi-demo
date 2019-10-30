package com.example.easypoidemo.export.test;


import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.export.styler.ExcelExportStylerDefaultImpl;
import com.example.easypoidemo.export.pojo.ScoreExcel;
import lombok.Cleanup;
import org.apache.commons.lang3.builder.ReflectionToStringBuilder;
import org.apache.poi.ss.usermodel.*;
import org.junit.Test;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * @Author: zhaoXinXing
 * @Date: 2019/10/24 14:01
 */
@Service
public class MyExportTest {
    private Random random = new Random();

    /**
     * 测试 普通导出
     *
     * @author zhaoXinXing
     * @date 2019/10/24
     */
    @Test
    public void test1111() {
        ArrayList<ScoreExcel> list = new ArrayList<>();
        for (int i = 1; i < 21; i++) {
            ScoreExcel scos = new ScoreExcel();
            scos.setAge(random.nextInt(7) + 11);
            scos.setSex(String.valueOf(random.nextInt(2)));
            scos.setName("小明" + i + "号");
            scos.setScore(random.nextInt(100));
            scos.setClassLevel(String.valueOf(random.nextInt(3)));
            scos.setLesson(String.valueOf(random.nextInt(12)));
            list.add(scos);
        }
        String title = "2019年期末学生成绩表";
        String fileName = "学生成绩表.xls";
        String sheetName = "表一";
        ExportParams exportParams = new ExportParams(title, sheetName, ExcelType.HSSF);
        exportParams.setStyle(ExcelExportStylerDefaultImpl.class);
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, ScoreExcel.class, list);
        try {
            @Cleanup FileOutputStream fos = new FileOutputStream(new File("D:/" + fileName));
            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 测试模板导出
     *
     * @author zhaoXinXing
     * @date 2019/10/24
     */
    @Test
    public void test1112() {
        Map<String, Object> map = new HashMap<String, Object>();
        Map<String, Object> modelMap = new HashMap<>();
        TemplateExportParams params = new TemplateExportParams("templates/myTemplate.xlsx");
        System.out.println(params.getTemplateUrl());
        System.out.println(params.getTemplateWb());
//        list数据
        List<ScoreExcel> listData = new ArrayList<ScoreExcel>();
        for (int i = 1; i < 10; i++) {
            ScoreExcel scos = new ScoreExcel();
            scos.setAge(random.nextInt(7) + 11);
            scos.setSex(String.valueOf(random.nextInt(2)));
            scos.setName("小明" + i + "号");
            scos.setScore(random.nextInt(100));
            scos.setClassLevel(String.valueOf(random.nextInt(3)));
            scos.setLesson(String.valueOf(random.nextInt(12)));
            listData.add(scos);
        }
        map.put("listData", listData);
        map.put("letest", "字符串长度是14应该显示14");
        map.put("fntest", "12345678.2341234");
        map.put("fdtest", new Date());
        List<Map<String, String>> mapList = new ArrayList<Map<String, String>>();
        for (int i = 0; i < 5; i++) {
            Map<String, String> testMap = new HashMap<String, String>();
            testMap.put("id", "xman" + i);
            testMap.put("name", "小明" + i);
            testMap.put("sex", "1");
            mapList.add(testMap);
        }
        map.put("maplist", mapList);

        mapList = new ArrayList<>();
        for (int i = 0; i < 6; i++) {
            Map<String, String> testMap = new HashMap<String, String>();
            testMap.put("xinde", "新的" + i);
            mapList.add(testMap);
        }
        map.put("sitest", mapList);
        Workbook workbook = ExcelExportUtil.exportExcel(params, map);
        try {
            @Cleanup FileOutputStream fos = new FileOutputStream(new File("D:/" + "模板导出.xlsx"));
            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 测试导入
     * 属性	             类型	         默认值	    功能
     * titleRows	     int	            0	    表格标题行数,默认0
     * headRows	         int	            1	    表头行数,默认1
     * startRows	     int	            0	    字段真正值和列标题之间的距离 默认0
     * keyIndex	         int	            0	    主键设置,如何这个cell没有值,就跳过 或者认为这个是list的下面的值,这一列必须有值,不然认为这列为无效数据
     * startSheetIndex	 int	            0	    开始读取的sheet位置,默认为0
     * sheetNum	         int	            1	    上传表格需要读取的sheet 数量,默认为1
     * needSave	        boolean	           false	    是否需要保存上传的Excel
     * needVerfiy	    boolean	           false	是否需要校验上传的Excel
     * saveUrl	        String	        "upload/excelUpload"	保存上传的Excel目录,默认是 如 TestEntity这个类保存路径就是upload/excelUpload/Test/yyyyMMddHHmss_ 保存名称上传时间_五位随机数
     * verifyHanlder   IExcelVerifyHandler	null	校验处理接口,自定义校验
     * lastOfInvalidRow	int	                0	    最后的无效行数,不读的行数
     * readRows	        int	                0	    手动控制读取的行数
     * importFields	    String[]	        null	导入时校验数据模板,是不是正确的Excel
     * keyMark	        String	            ":"	    Key-Value 读取标记,以这个为Key,后面一个Cell 为Value,多个改为ArrayList
     * readSingleCell	boolean	            false	按照Key-Value 规则读取全局扫描Excel,但是跳过List读取范围提升性能,仅仅支持titleRows + headRows + startRows 以及 lastOfInvalidRow
     * dataHanlder	   IExcelDataHandler	null	数据处理接口,以此为主,replace,format都在这后面
     *
     * @param null
     * @return
     * @author zhaoXinXing
     * @date 2019/10/28
     */
    @Test
    public void test1113() {
        ImportParams params = new ImportParams();
        long start = new Date().getTime();
        String resource = MyExportTest.class.getClassLoader().getResource("templates/importTemplate.xlsx").getPath();
//        System.out.println("resource = " + resource);
        //第一行表头不读取
        params.setHeadRows(1);
        //第二行标题行不读取
        params.setTitleRows(2);
        //读取前四行数据
//        params.setReadRows(4);

        //设置保存读取的excel，默认保存到"upload/excelUpload",会保存整个excel跟readRow无关
        params.setNeedSave(true);
//        params.setSaveUrl("templates/");

        //从第一个sheet开始，读取3个sheet
        params.setStartSheetIndex(0);
        params.setSheetNum(3);
        params.setLastOfInvalidRow(1);
        params.setNeedVerify(true);
        List<ScoreExcel> list = ExcelImportUtil.importExcel(new File(resource), ScoreExcel.class, params);
        System.out.println(new Date().getTime() - start);
        System.out.println(list.size());
        list.forEach(System.out::println);
//        list.forEach((a)-> System.out.println(a));
        System.out.println(ReflectionToStringBuilder.toString(list.get(0)));
    }

    /**
     * 使用poi导入
     *
     * @param null
     * @return
     * @author zhaoXinXing
     * @date 2019/10/29
     */
    @Test
    public void test1114() {
        String resource = MyExportTest.class.getClassLoader().getResource("templates/importTemplate.xlsx").getPath();
        int nameCl = 0;
        int ageCl = 0;
        int sexCl = 0;
        int scoreCl = 0;
        int lessonCl = 0;
        int classLevelCl = 0;
        int sumScoreCl = 0;
        int averageScoreCl = 0;
        List<ScoreExcel> list = new ArrayList<>();
        try {
            Workbook workbook = WorkbookFactory.create(new File(resource));
            for (Sheet sheet : workbook)
                for (Row row : sheet) {
                    if (row.getRowNum() == 2) {
                        for (Cell cell : row) {
                            switch (cell.getStringCellValue()) {
                                case ("姓名"):
                                    nameCl = cell.getColumnIndex();
                                case ("性别"):
                                    sexCl = cell.getColumnIndex();
                                case ("年龄"):
                                    ageCl = cell.getColumnIndex();
                                case ("分数"):
                                    scoreCl = cell.getColumnIndex();
                                case ("课程"):
                                    lessonCl = cell.getColumnIndex();
                                case ("年级"):
                                    classLevelCl = cell.getColumnIndex();
                                case ("总分"):
                                    sumScoreCl = cell.getColumnIndex();
                                case ("平均分"):
                                    averageScoreCl = cell.getColumnIndex();
                                default:
                            }
                        }
                    } else if (row.getRowNum() > 3 && row.getRowNum() < sheet.getLastRowNum()) {
                        ScoreExcel scoreExcel = new ScoreExcel();
                        for (Cell cell : row) {
                            int ci = cell.getColumnIndex();
                            if (ci == nameCl) {
                                scoreExcel.setName(cell.getStringCellValue());
                            } else if (ci == sexCl) {
                                scoreExcel.setSex(String.valueOf((int) cell.getNumericCellValue()));
                            } else if (ci == ageCl) {
                                scoreExcel.setAge((int) cell.getNumericCellValue());
                            } else if (scoreCl == ci) {
                                scoreExcel.setScore((int) cell.getNumericCellValue());
                            } else if (ci == lessonCl) {
                                scoreExcel.setLesson(String.valueOf((int) cell.getNumericCellValue()));
                            } else if (ci == classLevelCl) {
                                scoreExcel.setClassLevel(String.valueOf((int) cell.getNumericCellValue()));
                            } else if (ci == sumScoreCl) {
                                scoreExcel.setSumScore(cell.getNumericCellValue());
                            } else if (ci == averageScoreCl) {
                                scoreExcel.setAverageScore(cell.getNumericCellValue());
                            }
                        }
                        list.add(scoreExcel);
                    }
                }
        } catch (IOException e) {
            e.printStackTrace();
        }
        list.forEach(System.out::println);
    }
}
