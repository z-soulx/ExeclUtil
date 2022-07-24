/**
 * Created by d on 2018/9/27.
 */

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.beust.jcommander.internal.Lists;
import com.google.common.base.Charsets;
import com.google.common.io.Files;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.http.HttpEntity;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.conn.ssl.SSLConnectionSocketFactory;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClientBuilder;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.ssl.SSLContexts;
import org.apache.http.util.EntityUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;
import org.testng.collections.Sets;
import org.testng.internal.collections.Pair;

import javax.net.ssl.SSLContext;
import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.security.KeyManagementException;
import java.security.NoSuchAlgorithmException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

public class ExeclProcessT {

    @Test
    public void test() throws Exception {
        try {
            Pair<List<YQ>, List<YQ>> listListPair = doGet();
            List<YQ> dealyqs = dealyqs(listListPair);
            makeExcel(dealyqs);
        } catch (IOException | BiffException e) {
            e.printStackTrace();
            throw e;
        }
    }

    private static Set<String> readExcel(File file) throws IOException, BiffException {

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
        Set<String> allData = Sets.newHashSet();
        // 越过第一行 它是列名称
        for (int j = 2; j < rows; j++) {
            // 得到每一行的单元格的数据
            Cell[] cells = sheet.getRow(j);
            Cell cell = cells[4];
            allData.add(cell.getContents().replace("(新)","").trim());

        }
        return allData;

    }

    private List<YQ>  dealyqs(Pair<List<YQ>, List<YQ>> yqs) {
        Map<String,Integer> m = new HashMap<>();
        Map<String,Integer> g = new HashMap<>();
        List<YQ> gg = yqs.first().stream().flatMap(r -> {
            return r.getCommunitys().stream().map(d -> {
                YQ yq = new YQ();
                try {
                    BeanUtils.copyProperties(yq,r );
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                } catch (InvocationTargetException e) {
                    e.printStackTrace();
                }
                List<String> objects = Lists.newArrayList();
                String[] s = r.getArea_name().split(" ");
                objects.add(s[s.length - 1]+d);
                    g.put("c",g.getOrDefault("c",0) + 1);
                    g.put("S"+yq.getProvince(), g.getOrDefault("S"+yq.getProvince(), 0) + 1);
                    g.put("s"+yq.getCity(), g.getOrDefault("s"+yq.getCity(), 0) + 1);


                yq.setCommunitys(objects);
                return yq;
            });
        }).collect(Collectors.toList());

        List<YQ> mm = yqs.second().stream().flatMap(r -> {
            return r.getCommunitys().stream().map(d -> {
                YQ yq = new YQ();
                try {
                    BeanUtils.copyProperties(yq,r );
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                } catch (InvocationTargetException e) {
                    e.printStackTrace();
                }
                List<String> objects = Lists.newArrayList();
                String[] s = r.getArea_name().split(" ");
                objects.add(s[s.length - 1]+d);

                    m.put("c",m.getOrDefault("c",0) + 1);
                    m.put("S"+ yq.getProvince(), m.getOrDefault("S"+yq.getProvince(), 0) + 1);
                    m.put("s"+yq.getCity(), m.getOrDefault("s"+yq.getCity(), 0) + 1);


                yq.setCommunitys(objects);
                return yq;
            });
        }).collect(Collectors.toList());
        gg.sort(new Comparator<YQ>(){
            @Override
            public int compare(YQ o1, YQ o2) {
                return backGroup(o2) - backGroup(o1);
            }
        });
        mm.sort(new Comparator<YQ>(){
            @Override
            public int compare(YQ o1, YQ o2) {
                return backGroup(o2) - backGroup(o1);
            }
        });

        gg.forEach(r -> {
                r.setFx(r.getFx() + "(" + g.get("c") + ")");
                r.setProvince(r.getProvince() + "(" + g.get("S"+r.getProvince()) + ")");
                r.setCity(r.getCity() + "(" + g.get("s"+r.getCity()) + ")");

        });
        mm.forEach(r -> {
                r.setFx(r.getFx() + "(" + m.get("c") + ")");
                r.setProvince(r.getProvince() + "(" + m.get("S"+r.getProvince()) + ")");
                r.setCity(r.getCity() + "(" + m.get("s"+r.getCity()) + ")");
        });
        List<YQ> objects = Lists.newArrayList();
        objects.addAll(gg);
        objects.addAll(mm);
        return objects;
    }

    public  Pair<List<YQ>,List<YQ>>  doGet() throws IOException {
        //创建HttpClient对象
        CloseableHttpClient httpClient = HttpClientBuilder.create().build();
//        HttpGet get = new HttpGet("http://m.bendibao.com/news/gelizhengce/fengxianmingdan.php");
        HttpGet get = new HttpGet("https://covid-api.caduo.ml/latest.json?t="+System.currentTimeMillis());
//        Document doc = Jsoup.connect("http://m.bendibao.com/news/gelizhengce/fengxianmingdan.php").get();
        HttpEntity entity = httpClient.execute(get).getEntity();
        String s = EntityUtils.toString(entity);
//        String s = mock();
        JSONObject parse = JSON.parseObject(s);
        JSONArray highlist = parse.getJSONObject("data").getJSONArray("highlist");
        List<YQ> gyqs = highlist.toJavaList(YQ.class).stream().map(r-> {
            r.setFx("高风险区");
            return r;
        }).collect(Collectors.toList());
        JSONArray jsonArray2 = parse.getJSONObject("data").getJSONArray("middlelist");
        List<YQ> myqs = jsonArray2.toJavaList(YQ.class).stream().map(r-> {
            r.setFx("中风险区");
            return r;
        }).collect(Collectors.toList());;


        System.out.println();
        List<YQ> res = Lists.newArrayList();
        res.addAll(gyqs);
        res.addAll(myqs);
//        res.addAll(dyqs);
        return Pair.create(gyqs,myqs);
    }

public String mock() throws IOException {
    String result = Files.asCharSource(new File("E:\\aaa.txt"), Charsets.UTF_8).read();
    return result;
    }
    /**
     * 将数据写入到excel中
     */
    public static  void makeExcel( List<YQ>  result) throws IOException, BiffException {

        //第一步，创建一个workbook对应一个excel文件
        HSSFWorkbook workbook = new HSSFWorkbook();


        HSSFCellStyle fontStyle = workbook.createCellStyle();
        HSSFFont font = workbook.createFont();//创建字体
        font.setColor(IndexedColors.RED.getIndex());//设置字体颜色
        font.setFontName("方正楷体_GB方正楷体_GBKK");
        font.setFontHeightInPoints((short) 10);
        fontStyle.setFont(font);

        HSSFCellStyle fontcStyle = workbook.createCellStyle();
        HSSFFont font2 = workbook.createFont();//创建字体
        font2.setColor(IndexedColors.BLACK.getIndex());//设置字体颜色
        font2.setFontName("方正楷体_GB方正楷体_GBKK");
        font2.setFontHeightInPoints((short) 10);
        fontcStyle.setFont(font2);

        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        cellStyle.setFont(font2);

        HSSFCellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);


        HSSFCellStyle style2 = workbook.createCellStyle();
        style2.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style2.setFont(font);

        //第二部，在workbook中创建一个sheet对应excel中的sheet
        HSSFSheet sheet = workbook.createSheet("resultFormants");
        HSSFFont sfont = sheet.getWorkbook().getFontAt((short) 0);
        sfont.setColor(IndexedColors.BLACK.getIndex());//设置字体颜色
        sfont.setFontName("方正楷体_GB方正楷体_GBKK");
        sfont.setFontHeightInPoints((short) 10);
        //第三部，在sheet表中添加表头第0行，老版本的poi对sheet的行列有限制
        HSSFRow row = sheet.createRow(0);
        //第四步，创建单元格，设置表头
        HSSFCell cell = row.createCell(0);
        cell.setCellStyle(cellStyle);
        SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy 年 MM 月 dd 日 HH 时");
        String sd2 = sdf2.format(new Date(System.currentTimeMillis()));
        cell.setCellValue("全国疫情一览表（截至"+sd2+"）");
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, 4);
        sheet.addMergedRegion(region);

        HSSFRow row11 = sheet.createRow(1);
        row11.setRowStyle(cellStyle);
         row11.createCell(0).setCellValue("风险等级");
         row11.createCell(1).setCellValue("");
         row11.createCell(2).setCellValue("省（自治区、直辖市）");
         row11.createCell(3).setCellValue("城市（地区）");
         row11.createCell(4).setCellValue("具体地区");
        Set<String> strings = null;
        try {
            strings = readExcel(new File("E:\\全国中高风险区域一览表_220723.xls"));
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        }
        //第五步，写入数据
        int count1 = 1,count2 = 1;
        for (int i = 0; i < result.size(); i++) {
            YQ oneData = result.get(i);
            HSSFRow row1 = sheet.createRow(i + 2);

//            //创建单元格设值
            row1.createCell(0).setCellValue(oneData.getFx());
            HSSFCell cell1 = row1.createCell(1);
            cell1.setCellValue(oneData.getFx().contains("高") ? count1++ : count2++);
            HSSFCell cell2 = row1.createCell(2);
            cell2.setCellValue(oneData.getProvince());
            HSSFCell cell3 = row1.createCell(3);
            cell3.setCellValue(oneData.getCity());
            String s = oneData.getCommunitys().toString();
            String substring = s.substring(1, s.length() - 1);
            HSSFCell cell4 = row1.createCell(4);
            if (!strings.contains(substring)) {
                cell4.setCellStyle(fontStyle);
                cell4.setCellValue( substring+ "(新)");
            } else {
                cell4.setCellValue(s.substring(1,s.length() - 1));
            }

            if (oneData.getB()) {
                 cell2.setCellStyle(style);
                 cell3.setCellStyle(style);
                 cell4.setCellStyle(!strings.contains(substring) ? style2 : style);
            }


//
    }
        mergeSpecifiedColumn(sheet,0,workbook);
        mergeSpecifiedColumn(sheet,2, workbook);
        mergeSpecifiedColumn(sheet,3, workbook);

//        HSSFCellStyle g = workbook.createCellStyle();
//        style.setFillForegroundColor(IndexedColors.PINK.getIndex());
//        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//        sheet.getRow(2).getCell(0).setCellStyle(g);
//
//        HSSFCellStyle m = workbook.createCellStyle();
//        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
//        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//        String value = sheet.getRow(2).getCell(0).getStringCellValue();
//        String substring = value.substring(value.indexOf("(") + 1, value.indexOf(")"));
//        sheet.getRow(Integer.valueOf(substring) + 3).getCell(0).setCellStyle(m);

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

    private static int backGroup(YQ province) {
        String[] split = "新疆维吾尔自治区、西藏自治区、四川省、云南省、贵州省、重庆市、湖北省、陕西省、甘肃省、宁夏回族自治区、青海省".split("、");
        Set<String> ss = Sets.newHashSet();
        for (String s : split) {
            ss.add(s);
        }
        if (ss.contains(province.getProvince())) {
            province.setB(Boolean.TRUE);
            return 1;
        }
        return   0;
    }

    public static void mergeSpecifiedColumn(HSSFSheet sheet, int column, HSSFWorkbook workbook) {
        int totalRows = sheet.getLastRowNum(), firstRow = 0, lastRow = 0;
        boolean isLastCompareSame = false;//上一次比较是否相同
        //这里第一行是表头，从第三行开始判断是否相同
        if (totalRows >= 2 ) {
            for (int i = 2 ; i <= totalRows; i++) {
                String lastRowCellContent = sheet.getRow(i - 1).getCell(column).getStringCellValue();
                String curRowCellContent = sheet.getRow(i).getCell(column).getStringCellValue();
                if (curRowCellContent.equals(lastRowCellContent)) {
                    if (!isLastCompareSame) {
                        firstRow = i - 1;
                    }
                    lastRow = i;
                    isLastCompareSame = true;
                } else {
                    isLastCompareSame = false;
                    if (lastRow > firstRow) {
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow, lastRow, column, column);
                        sheet.addMergedRegion(cellRangeAddress);
//                        sheet.getRow(i - 1).getCell(column).setCellValue(lastRowCellContent +"("+(lastRow - firstRow)+")");
                        setRegionStyle(sheet,cellRangeAddress,workbook);
                        lastRow++;
                        firstRow = lastRow;
                    }
                }
                //最后一行时判断是否有需要合并的行
                if ((i == totalRows) && (lastRow > firstRow)) {
                    CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow, lastRow, column, column);
                    sheet.addMergedRegion(cellRangeAddress);
                    setRegionStyle(sheet,cellRangeAddress,workbook);
                }
            }
        }

    }

    private static void setRegionStyle(HSSFSheet sheet, CellRangeAddress region, HSSFWorkbook workbook) {

        HSSFCellStyle style = workbook.createCellStyle();
        HSSFCellStyle cs = style;
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        style.setBorderBottom(BorderStyle.THIN); //下边框
        style.setBorderLeft(BorderStyle.THIN);//左边框
        style.setBorderTop(BorderStyle.THIN);//上边框
        style.setBorderRight(BorderStyle.THIN);//右边框

        for (int i = region.getFirstRow(); i <= region.getLastRow(); i++) {
            HSSFRow row = sheet.getRow(i);
            HSSFCell cell = null;
            //循环设置单元格样式
            for (int j = region.getFirstColumn(); j <= region.getLastColumn(); j++) {
                cell = row.getCell((short) j);
                HSSFCellStyle cellStyle = cell.getCellStyle();
                cs.setFillForegroundColor(cellStyle.getFillForegroundColor());
                cs.setFillPattern(cellStyle.getFillPatternEnum());
                cs.setFont(cellStyle.getFont(workbook));
                cell.setCellStyle(cs);
            }
        }
    }


}