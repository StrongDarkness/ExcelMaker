package cn.qxl.bean;

import cn.qxl.annotation.Header;

/**
 * Created by qiu on 2019/1/4.
 */
public class testbean {
    @Header("字段a")
   private String a;
    @Header("字段b")
   private String b;
    @Header("字段c")
   private String c;

    public String getA() {
        return a;
    }

    public void setA(String a) {
        this.a = a;
    }

    public String getB() {
        return b;
    }

    public void setB(String b) {
        this.b = b;
    }

    public String getC() {
        return c;
    }

    public void setC(String c) {
        this.c = c;
    }
}
