import com.beust.jcommander.internal.Lists;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;
import jxl.read.biff.BiffException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.junit.Test;
import org.testng.collections.Maps;
import org.testng.collections.Sets;
import org.testng.internal.collections.Pair;

public class ExeclProcessBenDB {

    public static void main(String[] args) {
        System.out.println("Hello World!");
    }
    @Test
    public void test() throws Exception {
        try {
            Pair<List<YQ2>, Set<String>> source = source();
            makeExcel(source.first(),source.second());
        } catch (IOException | BiffException e) {
            e.printStackTrace();
            throw e;
        }
    }

    public Pair< List<YQ2>, Set<String>> source() throws IOException {

        Document doc = Jsoup.connect("http://m.bendibao.com/news/gelizhengce/fengxianmingdan.php").get();
        Elements select = doc.select("div.info");


        Elements g = select.select("div.height.info-item");
        String text = doc.select("div.fx-item.high-fx").text();
        String[] split = text.split("\\s");
        String s = "高风险(" + split[0].substring(0, split[0].length() - 1) + ")";
        List<YQ2> gyq = paseHtml(g, s,0);

        Elements m = select.select("div.middle.info-item");
        String text2 = doc.select("div.fx-item.middle-fx").text();
        String[] split2 = text2.split("\\s");
        String s2 = "中风险(" + split[0].substring(0, split2[0].length() - 1) + ")";
        List<YQ2> myq = paseHtml(m,s2 ,1);

        gyq.sort(new Comparator<YQ2>(){
            @Override
            public int compare(YQ2 o1, YQ2 o2) {
                return backGroup(o2) - backGroup(o1);
            }
        });
        myq.sort(new Comparator<YQ2>(){
            @Override
            public int compare(YQ2 o1, YQ2 o2) {
                return backGroup(o2) - backGroup(o1);
            }
        });

        gyq.forEach(r -> {
            r.setProvince(r.getProvince() +"("+countm.get(s+r.getProvince())+")");
        });

        myq.forEach(r -> {
            r.setProvince(r.getProvince() +"("+countm.get(s2+r.getProvince())+")");
        });

        gyq.addAll(myq);


        Elements md = select.select("div.tiaodi.info-item");
        List<YQ2> mdyq = paseHtml(md, "",5);
        Set<String> collect = mdyq.stream().map(r -> r.getCommunitys()).collect(Collectors.toSet());
        return Pair.create(gyq,collect);

    }
    private Map<String,Integer> countm = Maps.newHashMap();

    private List<YQ2> paseHtml(Elements g,String fx,int type) {
        List<YQ2> r = Lists.newArrayList();
        for (Element element1 : g) {
        for (Element element : element1.select("div.info-list")) {
            Elements select1 = element.select("div.shi.flex-between");
            Elements select2 = element.select("ul.info-detail");
            int size = select1.size();
            for (int i = 0; i < size; i++) {
                String[] split = select1.get(i).text().split("\\s");
                String shen = split[0];
                if(split.length == 3) {
                    split[2] = split[1];
                    split[1] = split[0];
                }
                String shi = split[1];
                Integer count = Integer.valueOf(split[2].substring(0,split[2].length() - 1));
                countm.put(fx+shen,countm.getOrDefault(fx+shen,0) + count);
                shi += "("+count+")";
//                System.out.println(shen + shi + count);

                for (Element element2 : select2.get(i).select("li")) {
                    String text = element2.text();
                    YQ2 y = new YQ2();
                    y.setProvince(shen);
                    y.setCity(shi);
                    y.setCommunitys(text);
                    y.setType(type);
                    y.setFx(fx);
                    r.add(y);

                }

            }

            System.out.println();

        }
        }
        return r;
    }


    private static int backGroup(YQ2 province) {
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



    /**
     * 将数据写入到excel中
     */
    public static  void makeExcel( List<YQ2>  result,Set<String> strings) throws IOException, BiffException {

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

        //第五步，写入数据
        int count1 = 1,count2 = 1;
        for (int i = 0; i < result.size(); i++) {
            YQ2 oneData = result.get(i);
            HSSFRow row1 = sheet.createRow(i + 2);

//            //创建单元格设值
            HSSFCell cell0 = row1.createCell(0);
            cell0.setCellValue(oneData.getFx());

            HSSFCellStyle ct = workbook.createCellStyle();
            ct.setFillForegroundColor(IndexedColors.PINK.getIndex());
            ct.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            ct.setVerticalAlignment(VerticalAlignment.CENTER);
            ct.setAlignment(HorizontalAlignment.CENTER_SELECTION);
            ct.setBorderBottom(BorderStyle.THIN); //下边框
            ct.setBorderLeft(BorderStyle.THIN);//左边框
            ct.setBorderTop(BorderStyle.THIN);//上边框
            ct.setBorderRight(BorderStyle.THIN);

            HSSFCellStyle ct2 = workbook.createCellStyle();
            ct2.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            ct2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            ct2.setVerticalAlignment(VerticalAlignment.CENTER);
            ct2.setAlignment(HorizontalAlignment.CENTER_SELECTION);
            ct2.setBorderBottom(BorderStyle.THIN); //下边框
            ct2.setBorderLeft(BorderStyle.THIN);//左边框
            ct2.setBorderTop(BorderStyle.THIN);//上边框
            ct2.setBorderRight(BorderStyle.THIN);

            cell0.setCellStyle(oneData.getType() == 0 ? ct : ct2);

            HSSFCell cell1 = row1.createCell(1);
            cell1.setCellValue(oneData.getType() == 0 ? count1++ : count2++);
            HSSFCell cell2 = row1.createCell(2);
            cell2.setCellValue(oneData.getProvince());
            HSSFCell cell3 = row1.createCell(3);
            cell3.setCellValue(oneData.getCity());
            String substring = oneData.getCommunitys();

            HSSFCell cell4 = row1.createCell(4);
            if (strings.contains(substring)) {
                cell4.setCellStyle(fontStyle);
                cell4.setCellValue( substring+ "(新)");
            } else {
                cell4.setCellValue(substring);
            }

            if (oneData.isB()) {
                cell2.setCellStyle(style);
                cell3.setCellStyle(style);
                cell4.setCellStyle(strings.contains(substring) ? style2 : style);
            }


//
        }
        mergeSpecifiedColumn(sheet,0,workbook);
        mergeSpecifiedColumn(sheet,2, workbook);
        mergeSpecifiedColumn(sheet,3, workbook);



        //将文件保存到指定的位置
        try {
            SimpleDateFormat sf = new SimpleDateFormat("yyMMdd");
            String sd = sf.format(new Date(System.currentTimeMillis()));

            File file = new File("/Users/soulx/Desktop/PG/msic/ExeclUtil/src/main/java/file/全国中高风险区域一览表_"+sd+".xls");
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
            for (int j = 2; j <= region.getLastColumn(); j++) {
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
