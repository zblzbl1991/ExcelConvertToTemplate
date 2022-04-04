package com.zbl;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.io.file.FileReader;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;

import java.io.File;
import java.math.BigDecimal;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelConvertTools {
    public static void main(String[] args) {
        export();
    }

    private static void export() {
        FileReader fileReader = new FileReader("20220404155348科目余额表_2022年第3期.xls");
        File file = fileReader.getFile();
        ExcelReader reader = ExcelUtil.getReader(file);
        System.out.println(file);
        List<List<Object>> read = reader.read(5);
        Map<String, BigDecimal> fillMap = new HashMap<>();
        for (List<Object> objects : read) {
            String key = String.valueOf(objects.get(0));
            String value = String.valueOf(objects.get(4));
            if (StrUtil.isNotEmpty(value)) {
                fillMap.put(key, new BigDecimal(value));
            }
        }
        //求设计薪酬
        getShejiXinchou(fillMap);
        //求工程薪酬
        getGongchengXinchou(fillMap);
        //求市场薪酬
        getShichangXinchou(fillMap);
        //求财务薪酬
        getCaiwuXinchou(fillMap);
        //求人力薪酬
        getRenliXinchou(fillMap);
        //求总经办薪酬
        getZongjingbanXinchou(fillMap);
        //求设计招待费用
        getShejiZhaodai(fillMap);
        //求工程招待费用
        getGongchengZhaodai(fillMap);
        //求市场招待
        getShichangZhaodai(fillMap);
        //求财务招待
        getCaiwuZhaodai(fillMap);
        //求人力招待
        getRenliZhaodai(fillMap);
        //求总经办招待
        getZongjingbanZhaodai(fillMap);
        //求固定费用合计
        getGudingHeji(fillMap);
        //求设计费用合计
        getShejiHeji(fillMap);
        //求工程合计
        getGongchengHeji(fillMap);
        //求市场合计
        getShichangHeji(fillMap);
        //求财务合计
        getCaiwuHeji(fillMap);
        //求人力合计
        getRenliHeji(fillMap);
        //求总经办合计
        getZongjingbanHeji(fillMap);
        FileReader temp = new FileReader("盈亏分析模板.xlsx");
        File tempFile = temp.getFile();
        ExcelWriter excelWriter = EasyExcel.write("导出数据" + DateUtil.today() + ".xlsx").withTemplate(tempFile).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        // 这里注意 入参用了forceNewRow 代表在写入list的时候不管list下面有没有空行 都会创建一行，然后下面的数据往后移动。默认 是false，会直接使用下一行，如果没有则创建。
        // forceNewRow 如果设置了true,有个缺点 就是他会把所有的数据都放到内存了，所以慎用
        // 简单的说 如果你的模板有list,且list不是最后一行，下面还有数据需要填充 就必须设置 forceNewRow=true 但是这个就会把所有数据放到内存 会很耗内存
        // 如果数据量大 list不是最后一行 参照下一个
        FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
        excelWriter.fill(fillMap, fillConfig, writeSheet);
        excelWriter.finish();
    }

    private static void getShejiZhaodai(Map<String, BigDecimal> fillMap) {
        BigDecimal gudingheji = fillMap.getOrDefault("54010109", BigDecimal.ZERO);
        fillMap.put("sheji_zhaodai", gudingheji);
    }
    private static void getGongchengZhaodai(Map<String, BigDecimal> fillMap) {
        BigDecimal gudingheji = fillMap.getOrDefault("54020209", BigDecimal.ZERO);
        fillMap.put("gongcheng_zhaodai", gudingheji);
    }
    private static void getShichangZhaodai(Map<String, BigDecimal> fillMap) {
        BigDecimal gudingheji = fillMap.getOrDefault("560107", BigDecimal.ZERO);
        fillMap.put("shichang_zhaodai", gudingheji);
    }
    private static void getCaiwuZhaodai(Map<String, BigDecimal> fillMap) {
        BigDecimal gudingheji = fillMap.getOrDefault("56022201_05", BigDecimal.ZERO).add(fillMap.getOrDefault("56022203_05", BigDecimal.ZERO)).add(fillMap.getOrDefault("56022204_05", BigDecimal.ZERO)).add(fillMap.getOrDefault("56022205_05", BigDecimal.ZERO));
        fillMap.put("caiwu_zhaodai", gudingheji);
    }
    private static void getRenliZhaodai(Map<String, BigDecimal> fillMap) {
        BigDecimal gudingheji = fillMap.getOrDefault("56022201_06", BigDecimal.ZERO).add(fillMap.getOrDefault("56022203_06", BigDecimal.ZERO)).add(fillMap.getOrDefault("56022204_06", BigDecimal.ZERO)).add(fillMap.getOrDefault("56022205_06", BigDecimal.ZERO));
        fillMap.put("renli_zhaodai", gudingheji);
    }
    private static void getZongjingbanZhaodai(Map<String, BigDecimal> fillMap) {
        BigDecimal gudingheji = fillMap.getOrDefault("56022201_01", BigDecimal.ZERO).add(fillMap.getOrDefault("56022203_01", BigDecimal.ZERO)).add(fillMap.getOrDefault("56022204_01", BigDecimal.ZERO)).add(fillMap.getOrDefault("56022205_01", BigDecimal.ZERO));
        fillMap.put("zongjingban_zhaodai", gudingheji);
    }


    private static void getGudingHeji(Map<String, BigDecimal> fillMap) {
        BigDecimal gudingheji = fillMap.getOrDefault("560204", BigDecimal.ZERO).add(fillMap.getOrDefault("560205", BigDecimal.ZERO)).add(fillMap.getOrDefault("560206", BigDecimal.ZERO)).add(fillMap.getOrDefault("560207", BigDecimal.ZERO)).add(fillMap.getOrDefault("560221", BigDecimal.ZERO)).add(fillMap.getOrDefault("560223", BigDecimal.ZERO)).add(fillMap.getOrDefault("560224", BigDecimal.ZERO)).add(fillMap.getOrDefault("560299", BigDecimal.ZERO));
        fillMap.put("guding_heji", gudingheji);
    }
    private static void getShejiHeji(Map<String, BigDecimal> fillMap) {
        BigDecimal gudingheji = fillMap.getOrDefault("sheji_xinchou", BigDecimal.ZERO).add(fillMap.getOrDefault("54010111", BigDecimal.ZERO)).add(fillMap.getOrDefault("54010105", BigDecimal.ZERO)).add(fillMap.getOrDefault("54010104", BigDecimal.ZERO)).add(fillMap.getOrDefault("54010108", BigDecimal.ZERO)).add(fillMap.getOrDefault("54010110", BigDecimal.ZERO)).add(fillMap.getOrDefault("54010112", BigDecimal.ZERO)).add(fillMap.getOrDefault("54010106", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("54010107", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("sheji_zhaodai", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("54010150", BigDecimal.ZERO))

                ;
        fillMap.put("sheji_heji", gudingheji);
    }
    private static void getShichangHeji(Map<String, BigDecimal> fillMap) {
        BigDecimal gudingheji = fillMap.getOrDefault("shichang_xinchou", BigDecimal.ZERO).add(fillMap.getOrDefault("560113", BigDecimal.ZERO)).add(fillMap.getOrDefault("560103", BigDecimal.ZERO)).add(fillMap.getOrDefault("560104", BigDecimal.ZERO)).add(fillMap.getOrDefault("560105", BigDecimal.ZERO)).add(fillMap.getOrDefault("560106", BigDecimal.ZERO)).add(fillMap.getOrDefault("560109", BigDecimal.ZERO)).add(fillMap.getOrDefault("560115", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("560114", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("560116", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("560118", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("560111", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("560108", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("560110", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("shichang_zhaodai", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("560199", BigDecimal.ZERO))

                ;
        fillMap.put("shichang_heji", gudingheji);
    }
    private static void getGongchengHeji(Map<String, BigDecimal> fillMap) {
        BigDecimal gudingheji = fillMap.getOrDefault("gongcheng_xinchou", BigDecimal.ZERO).add(fillMap.getOrDefault("5401020211", BigDecimal.ZERO)).add(fillMap.getOrDefault("5401020205", BigDecimal.ZERO)).add(fillMap.getOrDefault("5401020204", BigDecimal.ZERO)).add(fillMap.getOrDefault("54020208", BigDecimal.ZERO)).add(fillMap.getOrDefault("5401020210", BigDecimal.ZERO)).add(fillMap.getOrDefault("5401020212", BigDecimal.ZERO)).add(fillMap.getOrDefault("5401020206", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("5401020207", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("gongcheng_zhaodai", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("5401020250", BigDecimal.ZERO))

                ;
        fillMap.put("gongcheng_heji", gudingheji);
    }
    private static void getCaiwuHeji(Map<String, BigDecimal> fillMap) {
        BigDecimal gudingheji = fillMap.getOrDefault("caiwu_xinchou", BigDecimal.ZERO).add(fillMap.getOrDefault("560203_05", BigDecimal.ZERO)).add(fillMap.getOrDefault("560208_05", BigDecimal.ZERO)).add(fillMap.getOrDefault("560209_05", BigDecimal.ZERO)).add(fillMap.getOrDefault("560210_05", BigDecimal.ZERO)).add(fillMap.getOrDefault("560214_05", BigDecimal.ZERO)).add(fillMap.getOrDefault("560215_05", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("560216_05", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("560219_05", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("caiwu_zhaodai", BigDecimal.ZERO))

                ;
        fillMap.put("caiwu_heji", gudingheji);
    }
    private static void getRenliHeji(Map<String, BigDecimal> fillMap) {
        BigDecimal gudingheji = fillMap.getOrDefault("renli_xinchou", BigDecimal.ZERO).add(fillMap.getOrDefault("560203_06", BigDecimal.ZERO)).add(fillMap.getOrDefault("560208_06", BigDecimal.ZERO)).add(fillMap.getOrDefault("560209_06", BigDecimal.ZERO)).add(fillMap.getOrDefault("560210_06", BigDecimal.ZERO)).add(fillMap.getOrDefault("560214_06", BigDecimal.ZERO)).add(fillMap.getOrDefault("560215_06", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("560216_06", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("560219_06", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("renli_zhaodai", BigDecimal.ZERO))

                ;
        fillMap.put("renli_heji", gudingheji);
    }
    private static void getZongjingbanHeji(Map<String, BigDecimal> fillMap) {
        BigDecimal gudingheji = fillMap.getOrDefault("zongjingban_xinchou", BigDecimal.ZERO).add(fillMap.getOrDefault("560203_01", BigDecimal.ZERO)).add(fillMap.getOrDefault("560208_01", BigDecimal.ZERO)).add(fillMap.getOrDefault("560209_01", BigDecimal.ZERO)).add(fillMap.getOrDefault("560210_01", BigDecimal.ZERO)).add(fillMap.getOrDefault("560214_01", BigDecimal.ZERO)).add(fillMap.getOrDefault("560215_01", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("560216_01", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("560219_01", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("560202", BigDecimal.ZERO))
                .add(fillMap.getOrDefault("zongjingban_zhaodai", BigDecimal.ZERO))

                ;
        fillMap.put("zongjingban_heji", gudingheji);
    }

    private static void getShejiXinchou(Map<String, BigDecimal> fillMap) {
        BigDecimal shejiXinchou = fillMap.getOrDefault("54010101", BigDecimal.ZERO).add(fillMap.getOrDefault("54010102", BigDecimal.ZERO)).add(fillMap.getOrDefault("54010103", BigDecimal.ZERO));
        fillMap.put("sheji_xinchou", shejiXinchou);
    }

    private static void getGongchengXinchou(Map<String, BigDecimal> fillMap) {
        BigDecimal shejiXinchou = fillMap.getOrDefault("5401020201", BigDecimal.ZERO).add(fillMap.getOrDefault("5401020202", BigDecimal.ZERO)).add(fillMap.getOrDefault("5401020203", BigDecimal.ZERO));
        fillMap.put("gongcheng_xinchou", shejiXinchou);
    }
    private static void getShichangXinchou(Map<String, BigDecimal> fillMap) {
        BigDecimal shejiXinchou = fillMap.getOrDefault("560101", BigDecimal.ZERO).add(fillMap.getOrDefault("560102", BigDecimal.ZERO)).add(fillMap.getOrDefault("560117", BigDecimal.ZERO));
        fillMap.put("shichang_xinchou", shejiXinchou);
    }
    private static void getCaiwuXinchou(Map<String, BigDecimal> fillMap) {
        BigDecimal shejiXinchou = fillMap.getOrDefault("56020101_05", BigDecimal.ZERO).add(fillMap.getOrDefault("56020102_05", BigDecimal.ZERO)).add(fillMap.getOrDefault("56020103_05", BigDecimal.ZERO));
        fillMap.put("caiwu_xinchou", shejiXinchou);
    }
    private static void getRenliXinchou(Map<String, BigDecimal> fillMap) {
        BigDecimal shejiXinchou = fillMap.getOrDefault("56020101_06", BigDecimal.ZERO).add(fillMap.getOrDefault("56020102_06", BigDecimal.ZERO)).add(fillMap.getOrDefault("56020103_06", BigDecimal.ZERO));
        fillMap.put("renli_xinchou", shejiXinchou);
    }
    private static void getZongjingbanXinchou(Map<String, BigDecimal> fillMap) {
        BigDecimal shejiXinchou = fillMap.getOrDefault("56020101_01", BigDecimal.ZERO).add(fillMap.getOrDefault("56020102_01", BigDecimal.ZERO)).add(fillMap.getOrDefault("56020103_01", BigDecimal.ZERO));
        fillMap.put("zongjingban_xinchou", shejiXinchou);
    }
}
