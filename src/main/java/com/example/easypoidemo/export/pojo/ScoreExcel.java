package com.example.easypoidemo.export.pojo;

import cn.afterturn.easypoi.excel.annotation.Excel;
import com.example.easypoidemo.export.Utils.BeanDealUtils;
import lombok.Data;

import java.io.Serializable;

/**
 * 模板导出测试类
 * @Author: zhaoXinXing
 * @Date: 2019/10/24 14:02
 */
@Data
public class ScoreExcel implements Serializable {
    @Excel(name="姓名")
    private String name;
    @Excel(name="性别",replace = {"男_1","女_0"})
    private String sex;
    @Excel(name="年龄")
    private Integer age;
    @Excel(name="分数")
    private Integer score;
    @Excel(name="课程",replace = {"语文_0","数学_1","英语_2","地理_3","历史_4","政治_5","生物_6"})
    private String lesson;
    @Excel(name="年级",replace = {"七年级_0","八年级_1","九年级_2"})
    private String classLevel;
    @Excel(name="总分")
    private Double sumScore;
    @Excel(name="平均分")
    private Double averageScore;


    public String getSex() {
        return BeanDealUtils.getValueByKey(this.sex,SEX.class);
    }

    public String getLesson() {
        return BeanDealUtils.getValueByKey(this.lesson,LESSONS.class);
    }

    public String getClassLevel() {
        return BeanDealUtils.getValueByKey(this.classLevel,CLASSLEVEL.class);
    }

    public enum SEX{
        MAN("1","男"),WOMEN("0","女");
        private String name;
        private String value;

        SEX(String value, String name) {
            this.name = name;
            this.value = value;
        }

        public static String getNameByValue(String value) {
            for (SEX sex : SEX.values()) {
                if(sex.value.equals(value)){
                    return sex.name;
                }
            }
            return null;
        }

        public String getValue() {
            return value;
        }
    }
    public enum CLASSLEVEL{
        SEVEN("0","七年级"),
        EIGHT("1","八年级"),
        NINE("2","九年级");
        private String name;
        private String value;

        CLASSLEVEL(String value, String name) {
            this.name = name;
            this.value = value;
        }

        public static String getNameByValue(String value) {
            for (CLASSLEVEL cl : CLASSLEVEL.values()) {
                if(cl.value.equals(value)){
                    return cl.name;
                }
            }
            return null;
        }

        public String getValue() {
            return value;
        }
    }

    public enum LESSONS{
        CHINESE("0","语文"),
        MATHS("1","数学"),
        ENGLISH("2","英语"),
        PHYSICS("3","物理"),
        BIOLOGY("4","生物"),
        POLITICS("5","政治"),
        HISTORY("6","历史"),
        GEOGRAPHY("7","地理"),
        PE("8","体育"),
        ART("9","美术"),
        MUSIC("10","音乐"),
        CHEMISTRY("11","化学");
        private String name;
        private String value;

        LESSONS(String value, String name) {
            this.name = name;
            this.value = value;
        }

        public static String getNameByValue(String value) {
            for (LESSONS les : LESSONS.values()) {
                if(les.value.equals(value)){
                    return les.name;
                }
            }
            return null;
        }

        public String getValue() {
            return value;
        }
    }

}
