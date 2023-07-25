package com.example.config;

import com.deepoove.poi.XWPFTemplate;
import org.junit.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class WordTest {

    @Test
    public void test2() throws IOException {
        // word模板的绝对路径
        String templatePath = "E:\\workspace1\\filemaster\\file-master\\template.docx";
        String outPath = "E:\\workspace1\\filemaster\\file-master\\templateTest.docx";
        Map<String, Object> map = new HashMap<>();

        // 获取当前时间的年、月和日
        SimpleDateFormat yearFormat = new SimpleDateFormat("yyyy");
        SimpleDateFormat monthFormat = new SimpleDateFormat("MM");
        SimpleDateFormat dayFormat = new SimpleDateFormat("dd");
        Date currentDate = new Date();
        String year = yearFormat.format(currentDate);
        String month = monthFormat.format(currentDate);
        String day = dayFormat.format(currentDate);

        // 设置模板中的日期信息
        map.put("year", year);//
        map.put("month", month);//
        map.put("day", day);//
        map.put("fuName", "ピカチュウ");//
        map.put("name", "皮卡丘");//
        map.put("age", "25");//
        map.put("gender", "男");//
        map.put("fuAddress", "とうきょうとちよだくそとがいかんだ５ちょうめ１−２ すえひろビル ５F");//
        map.put("post", "101-0021");//
        map.put("address", "東京都千代田区外神田５丁目１−２ 末広ビル 5F");//
        map.put("phone", "080 1234 1234");//
        map.put("motivation", "ピカピカピカピカピカピカピカピカピカピカピカピカピカピカピカピカピカピカ");//
        map.put("hope", "ピカチュウ、ピカチュウピカチュウ、ピカチュウピカチュウ、ピカチュウピカチュウ、ピカチュウ");//

        XWPFTemplate template = XWPFTemplate.compile(templatePath).render(map);
        template.writeAndClose(Files.newOutputStream(Paths.get(outPath)));
    }
}
