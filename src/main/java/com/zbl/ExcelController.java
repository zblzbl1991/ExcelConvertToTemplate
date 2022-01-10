package com.zbl;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.io.file.FileReader;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * ClassName: ExcelController
 * Description: <br/>
 * date: 2022/1/10 14:45
 *
 * @author 赵宝龙
 * @since JDK 1.8
 */
@Controller
public class ExcelController {
    @GetMapping("/")
    public String index(){
        return "index";
    }
    /**
     * 实现文件上传
     * */
    @RequestMapping(value = "fileUpload",produces = "application/octet-stream")
    public void fileUpload(@RequestParam("fileName") MultipartFile file,
                           HttpServletResponse response) throws IOException {
        if(file.isEmpty()){
           throw new RuntimeException();
        }
        InputStream inputStream = file.getInputStream();
        ExcelReader reader = ExcelUtil.getReader(inputStream);
        System.out.println(file);
        List<List<Object>> read = reader.read(5);
        Map<String,Object> fillMap =new HashMap<>();
        for (List<Object> objects : read) {
            String key = String.valueOf(objects.get(0)) ;
            Object value = objects.get(8);
            fillMap.put(key,value);
        }
        FileReader temp = new FileReader("盈亏分析模板.xlsx");
        File tempFile = temp.getFile();
        String fileName = "导出数据" + DateUtil.today() + ".xlsx";
        String encode = URLEncoder.encode(fileName, "UTF-8");
        response.setHeader("Content-Disposition","attachment;filename="+encode);
        response.setContentType("application/vnd.ms-excel;charset=UTF-8");
        response.setHeader("Access-Control-Expose-Headers","Content-Disposition");
        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream()).withTemplate(tempFile).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        // 这里注意 入参用了forceNewRow 代表在写入list的时候不管list下面有没有空行 都会创建一行，然后下面的数据往后移动。默认 是false，会直接使用下一行，如果没有则创建。
        // forceNewRow 如果设置了true,有个缺点 就是他会把所有的数据都放到内存了，所以慎用
        // 简单的说 如果你的模板有list,且list不是最后一行，下面还有数据需要填充 就必须设置 forceNewRow=true 但是这个就会把所有数据放到内存 会很耗内存
        // 如果数据量大 list不是最后一行 参照下一个
        FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
        excelWriter.fill(fillMap,fillConfig, writeSheet);
        excelWriter.finish();
        }

}
