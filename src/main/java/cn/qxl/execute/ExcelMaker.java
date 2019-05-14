package cn.qxl.execute;

import cn.qxl.annotation.*;
import cn.qxl.bean.testbean;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 导出Excel工具类(通用)
 *
 * @author qiu
 */
public class ExcelMaker<T> {
    private String title;
    private List<String> header;
    private List<T> data;
    // 第一步，创建一个webbook，对应一个Excel文件
    private static HSSFWorkbook wb = new HSSFWorkbook();

    public ExcelMaker() {
    }

    public ExcelMaker(String title, List<String> header, List<T> data) {
        this.title = title;
        this.header = header;
        this.data = data;
    }


    public ExcelMaker<T> create() {
        if (title == null) {
            throw new IllegalArgumentException("表名称不能为空！");
        }
        // 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
        int pageSize = 50000;
        int index = data.size() / pageSize;// 记录额外创建的sheet数量
        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        // style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式
        for (int i = 1; i < index + 2; i++) {
            HSSFSheet sheet = wb.createSheet(title + i);
            HSSFRow row = sheet.createRow((int) 0);
            for (int j = 0; j < header.size(); j++) {
                HSSFCell cell = row.createCell(j);
                cell.setCellValue(header.get(j));
            }
            // System.out.println(title);
            if (data == null||data.size()==0) {
                return this;
            }
            if (data.get(0) instanceof Map) {
                if (header==null){
                    throw new IllegalArgumentException("表头不能为空！");
                }
                // 超过50000行分页
                for (int n = i * pageSize - pageSize + 1; n < i * pageSize && n < data.size() + 1; n++) {
                    row = sheet.createRow((int) n - pageSize * i + pageSize);
                    Map map = (Map) data.get(i);
                    int j = 0;
                    for (Object o : map.keySet()) {
                        Object v = map.get(o);
                        if (v == null) {
                            row.createCell(j).setCellValue("");
                        } else if (v instanceof Integer) {
                            row.createCell(j).setCellValue((Integer) v);
                        } else if (v instanceof Date) {
                            SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                            row.createCell(j).setCellValue(format.format(v));
                        } else if (v instanceof Double) {
                            row.createCell(j).setCellValue((double) v);
                        } else {
                            row.createCell(j).setCellValue(v.toString());
                        }
                        j++;
                    }
                }
            } else if (data.get(0) instanceof Object) {
                Class c = data.get(0).getClass();
                Field[] fields = c.getDeclaredFields();
                for (int f = 0; f < fields.length; f++) {
                    Field field = fields[f];
                    //表字段名
                    if (field.isAnnotationPresent(Header.class)){
                        Header header=field.getAnnotation(Header.class);
//                        HSSFRow headRow = sheet.createRow(0);
                        HSSFCell hcell = row.createCell(f);
                        hcell.setCellValue(header.value());
                    }
                }
                // 超过50000行分页
                for (int n = i * pageSize - pageSize + 1; n < i * pageSize && n < data.size()+i ; n++) {
                    row = sheet.createRow((int)( n - pageSize * i + pageSize));
                    int step = 0;//忽略字段跳过
                    outer:
                    for (int j = 0; j < fields.length; j++) {
                        Field field = fields[j];
                        //字段方法名
                        String methodName = "get" + field.getName().substring(0, 1).toUpperCase()
                                + field.getName().substring(1);
                        try {
                            Method method = c.getMethod(methodName, new Class[]{});
                            Object value = method.invoke(data.get(n-i), new Object[]{});
                            // System.out.println(value);
                            Annotation[] anns = field.getAnnotations();
                            for (Annotation ann : anns) {
                                if (field.isAnnotationPresent(IngnoreField.class)) {
                                    // System.out.println(field.getName());
                                    step++;
                                    continue outer;// 跳出指定循环
                                }
                                if (field.isAnnotationPresent(StateFields.class)) {
                                    StateFields sfs = field.getAnnotation(StateFields.class);
                                    for (StateField sf : sfs.value()) {
                                        if (value instanceof String) {
                                            if (value.toString().trim().equals(sf.key())) {
                                                // System.out.println(sf.value());
                                                value = sf.value();
                                            }
                                        }
                                    }
                                } else if (field.isAnnotationPresent(StateIntFields.class)) {
                                    StateIntFields sifs = field.getAnnotation(StateIntFields.class);
                                    for (StateIntField sif : sifs.value()) {
                                        if (value instanceof Integer) {
                                            // System.out.println(value + "-" +
                                            // sif.key()+"-"+sif.value());
                                            if (Integer.parseInt(value + "") == sif.key()) {
                                                value = sif.value();
                                            }
                                        }
                                    }
                                }
                            }
                            if (value == null) {
                                row.createCell(j - step).setCellValue("");
                            } else {
                                if (value instanceof Integer) {
                                    row.createCell(j - step).setCellValue((Integer) value);
                                } else if (value instanceof Date) {
                                    SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                                    row.createCell(j - step).setCellValue(format.format(value));
                                } else if (value instanceof Double) {
                                    row.createCell(j - step).setCellValue((double) value);
                                } else {
                                    row.createCell(j - step).setCellValue(value.toString());
                                }
                            }
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                        row.setRowStyle(style);
                    }
                }

            } else {
                throw new IllegalArgumentException("数据类型只支持Object或者map类型！");
            }
        }
        // for (int i = 0; i < header.size(); i++) {
        // sheet.autoSizeColumn((short) i, true);// 自动列宽
        // }
        return this;
    }

    public static class Builder<T> {
        String title;
        List<String> headers = new ArrayList<String>();
        List<T> data;

        /**
         * 表格名称
         *
         * @param title
         * @return
         */
        public Builder<T> setTitle(String title) {
            this.title = title;
            return this;
        }

        /**
         * 添加表头
         *
         * @param header
         * @return
         */
        public Builder<T> addHeader(String header) {
            headers.add(header);
            return this;
        }

        /**
         * 添加数据
         *
         * @param data
         * @return
         */
        public Builder<T> setData(List<T> data) {
            this.data = data;
            return this;
        }

        /**
         * 生成表格
         *
         * @return
         */
        public ExcelMaker<T> build() {
            return new ExcelMaker<T>(title, headers, data);
        }
    }

    /**
     * 获取文件
     *
     * @param fileName
     * @return
     */
    public static File getFile(String fileName) {
        File file = null;
        try {
            file = new File(fileName);
            FileOutputStream fos = new FileOutputStream(file);
            wb.write(fos);
            fos.flush();
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return file;
    }

    public static void main(String[] args) {

        // ExcelMaker maker = new
        // ExcelMaker.Builder<Orders>().setTitle("标题").addHeader("表头1").addHeader("表头2").setData(ol)
        // .build();
        List<Map<String, Object>> list2 = new ArrayList<Map<String, Object>>();
        Map<String, Object> map = new LinkedHashMap<String, Object>();
        map.put("a", "1");
        map.put("b", "2");
        map.put("c", "3");
        map.put("d", "4");
        map.put("e", "5");
        map.put("f", "6");
        map.put("g", new Date());
        for (int i = 0; i < 70000; i++) {
            list2.add(map);
        }
        List<testbean> list = new ArrayList<>();
        testbean tb = new testbean();
        tb.setA("a");
        tb.setB("b");
        tb.setC("c");
        for (int i = 0; i < 70000; i++) {
            list.add(tb);
        }
//        ExcelMaker.Builder<testbean> builder = new ExcelMaker.Builder<testbean>();
////        ExcelMaker.Builder<Map<String, Object>> builder = new ExcelMaker.Builder<Map<String, Object>>();
//        builder.setTitle("标题");
//        builder.addHeader("标题1");
//        builder.addHeader("标题2");
//        builder.addHeader("标题3");
//        builder.addHeader("标题4");
//        builder.addHeader("标题5");
//        builder.addHeader("标题6");
//        builder.addHeader("标题7");
//        builder.setData(list);
//        ExcelMaker maker = builder.build().create();
//        HSSFWorkbook wb = maker.create();
        // String[] head=new String[5];
        // for (int i = 0; i < head.length; i++) {
        // head[i]="表头"+i;
        //// }
        // HSSFWorkbook wb =createExcel("标题", head, ol);
        File f = new File("E://excel");
        if (!f.exists()) {
            f.mkdir();
        }
//        try {
//            FileOutputStream fos = new FileOutputStream(f + "//1.xls");
//            wb.write(fos);
//            // fos.flush();
//            fos.close();
//        maker.getFile("E://excel//1.xls");
//        new ExcelMaker.Builder<Map<String, Object>>().setTitle("表").addHeader("1").addHeader("2").addHeader("3").setData(list2).build().create().getFile("E://excel//1.xls");
        new ExcelMaker.Builder<testbean>().setTitle("表").addHeader("1").addHeader("2").addHeader("3").setData(list).build().create().getFile("E://excel//2.xls");
        System.out.println("生成成功");
//        } catch (FileNotFoundException e) {
//            e.printStackTrace();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
    }

}
