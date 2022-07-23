/**
 * Created by d on 2018/9/27.
 */

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.google.common.collect.ArrayListMultimap;
import com.google.common.collect.Multimap;
import jdk.nashorn.internal.runtime.JSONFunctions;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.HttpStatus;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClientBuilder;
import org.apache.http.util.EntityUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.junit.Test;

import java.io.*;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

public class ExeclProcessT {

    public static void main(String[] args) {

        // 读取Excel文件
//        File file = new File("D:/execl.xls");
        try {
            //得到所有数据
            List<Bo> allData = read();

            //直接将它写到excel中
//            List<Bo> result = dealData(allData);
            List<Bo> result = null;
//            makeExcel(result);

        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

    }

    private static List<Bo> read() {

        return null;
    }
    @Test
    public void test() {
        try {
            List<YQ> yqs = doGet();
            makeExcel(yqs);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public  List<YQ>  doGet() throws IOException {
        //创建HttpClient对象
        CloseableHttpClient httpClient = HttpClientBuilder.create().build();
//        HttpGet get = new HttpGet("http://m.bendibao.com/news/gelizhengce/fengxianmingdan.php");
        HttpGet get = new HttpGet("https://covid-api.caduo.ml/latest.json?t=1658580290725");
//        Document doc = Jsoup.connect("http://m.bendibao.com/news/gelizhengce/fengxianmingdan.php").get();
        HttpEntity entity = httpClient.execute(get).getEntity();
        String s = EntityUtils.toString(entity);
        JSONObject parse = JSON.parseObject(s);
        JSONArray highlist = parse.getJSONObject("data").getJSONArray("highlist");
        List<YQ> yqs = highlist.toJavaList(YQ.class);

        System.out.println();

        return yqs;
    }


    /**
     * 获取数据
     * @param file
     * @return
     * @throws Exception
     */
    private static List<Bo> readExcel(File file) throws Exception {

        // 创建输入流，读取Excel
        InputStream is = new FileInputStream(file.getAbsolutePath());
        // jxl提供的Workbook类
        Workbook wb = Workbook.getWorkbook(is);
        // 只有一个sheet,直接处理
        //创建一个Sheet对象
        Sheet sheet = wb.getSheet(0);
        // 得到所有的行数
        int rows = sheet.getRows();
        // 所有的数据
        List<Bo> allData = new ArrayList();
        // 越过第一行 它是列名称
        for (int j = 1; j < rows; j++) {

            // 得到每一行的单元格的数据
            Cell[] cells = sheet.getRow(j);
            Bo oneData = new Bo();
            oneData.setItem(cells[0].getContents().trim());
            oneData.setSubject(cells[1].getContents().trim());
            oneData.setVowel(cells[2].getContents().trim());
            oneData.setNode(cells[3].getContents().trim());
            oneData.setF1(Integer.parseInt(cells[4].getContents().trim()));
            oneData.setF2(Integer.parseInt(cells[5].getContents().trim()));
            oneData.setF3(Integer.parseInt(cells[6].getContents().trim()));
            oneData.setF4(Integer.parseInt(cells[7].getContents().trim()));
            // 存储每一条数据
            allData.add(oneData);
            // 打印出每一条数据
            //System.out.println(oneData);

        }
        return allData;

    }

    /**
     * 处理数据
     */
    public static  List<Bo> dealData(List<Bo> allData) {
        Multimap<String, Bo> dataMap = ArrayListMultimap.create();        //结果
        List<Bo> result=new ArrayList();
        DecimalFormat fnum1 = new DecimalFormat("##0.00");
       for(int i=0;i<allData.size();i++) {
           Bo oneData = allData.get(i);
            String key = oneData.getSubject()+oneData.getVowel()+oneData.getNode();
            dataMap.put(key,oneData);
        }
        for(String key:dataMap.keySet())
        {
            List<Bo> dataList = (List<Bo>) dataMap.get(key);
            Bo avData = new Bo();
            double f1 = 0;
            double f2 = 0;
            double f3 = 0;
            double f4 = 0;
            Integer size = dataList.size();
            for(Bo vo:dataList)
            {
                avData.setSubject(vo.getSubject());
                avData.setVowel(vo.getVowel());
                avData.setNode(vo.getNode());
                f1 = f1+vo.getF1();
                f2 = f2+vo.getF2();
                f3 = f3+vo.getF3();
                f4 = f4+vo.getF4();
            }
            double avF1 = f1/size;
            double avF2 = f2/size;
            double avF3 = f3/size;
            double avF4 = f4/size;
            avData.setF1(Double.valueOf(fnum1.format(avF1)));
            avData.setF2(Double.valueOf(fnum1.format(avF2)));
            avData.setF3(Double.valueOf(fnum1.format(avF3)));
            avData.setF4(Double.valueOf(fnum1.format(avF4)));
            result.add(avData);
        }

        return result;
    }

    /**
     * 将数据写入到excel中
     */
    public static  void makeExcel( List<YQ>  result) {

        //第一步，创建一个workbook对应一个excel文件
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        //第二部，在workbook中创建一个sheet对应excel中的sheet
        HSSFSheet sheet = workbook.createSheet("resultFormants");
        //第三部，在sheet表中添加表头第0行，老版本的poi对sheet的行列有限制
        HSSFRow row = sheet.createRow(0);
        //第四步，创建单元格，设置表头
        HSSFCell cell = row.createCell(0);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("全国疫情一览表（截至2022年7月23日17时）");
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, 4);
        sheet.addMergedRegion(region);

        HSSFRow row11 = sheet.createRow(1);
        row11.setRowStyle(cellStyle);
         row11.createCell(0).setCellValue("风险等级");
         row11.createCell(1).setCellValue("");
         row11.createCell(2).setCellValue("省（自治区、直辖市）");
         row11.createCell(3).setCellValue("城市（地区）");
         row11.createCell(4).setCellValue("具体地区");

        //第五步，写入数据
        for (int i = 0; i < result.size(); i++) {
            YQ oneData = result.get(i);
            HSSFRow row1 = sheet.createRow(i + 1);
//            //创建单元格设值
            row1.createCell(0).setCellValue(oneData.getProvince());
            row1.createCell(1).setCellValue(oneData.getCity());
            row1.createCell(2).setCellValue(oneData.getArea_name());
            row1.createCell(3).setCellValue(oneData.getCommunitys().toString());
//            row1.createCell(4).setCellValue(oneData.getF2());
//            row1.createCell(5).setCellValue(oneData.getF3());
//            row1.createCell(6).setCellValue(oneData.getF4());
    }

        //将文件保存到指定的位置
        try {
            File file = new File("E:\\result.xls");
            if (file.exists()) {
                file.delete();
            }
            FileOutputStream fos = new FileOutputStream(file);

            workbook.write(fos);
            System.out.println("写入成功");
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


}