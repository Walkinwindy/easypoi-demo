package com.example.easypoidemo.export.Utils;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 * 对象处理类
 *
 * @author zhaoXinXing
 * @date 2019/10/28 11:46
 */
public class BeanDealUtils {
  /**
   * 通过代号获取名称
   * @author zhaoXinXing
   * @date 2019/10/28
   */
  public static String getValueByKey(String value, Class cls) {
    Method getNameByValue = null;
    String invoke = null;
    try {
      getNameByValue = cls.getMethod("getNameByValue", String.class);
    invoke = (String)getNameByValue.invoke(null, value);
    }catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException e) {
      invoke = "参数错误";
  }
    return invoke;
  }
}
