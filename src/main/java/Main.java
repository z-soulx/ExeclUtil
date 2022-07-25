import com.beust.jcommander.internal.Lists;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.junit.Test;
import org.testng.collections.Maps;

public class Main {

    public static void main(String[] args) {
        System.out.println("Hello World!");
    }

    @Test
    public void test() throws IOException {
        Map<String, Map<String,String>> gmap = Maps.newHashMap();
        Map<String, Map<String,String>> mmap = Maps.newHashMap();
        List<YQ2> all = Lists.newArrayList();
        Document doc = Jsoup.connect("http://m.bendibao.com/news/gelizhengce/fengxianmingdan.php").get();
        Elements select = doc.select("div.info");

//        Elements div = select.select("div.flex-between.bold");
        Elements g = select.select("div.height.info-item");

//        Elements div2 = select.select("ul.info-detail");
        List<YQ2> gyq = paseHtml(g, 0);
        Elements m = select.select("div.middle.info-item");
        List<YQ2> myq = paseHtml(m, 1);
        gyq.addAll(myq);
        gyq.forEach(r -> {
            r.setProvince(r.getProvince() +"("+countm.get(r.getProvince())+")");
        });

        Elements md = select.select("div.tiaodi.info-item");
        List<YQ2> mdyq = paseHtml(md, 5);
        Set<String> collect = mdyq.stream().map(r -> r.getCommunitys()).collect(Collectors.toSet());
        System.out.println();



    }
    private Map<String,Integer> countm = Maps.newHashMap();

    private List<YQ2> paseHtml(Elements g,int type) {
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
                countm.put(shen,countm.getOrDefault(shen,0) + count);
                shi += "("+count+")";
                System.out.println(shen + shi + count);

                for (Element element2 : select2.get(i).select("li")) {
                    String text = element2.text();
                    YQ2 y = new YQ2();
                    y.setProvince(shen);
                    y.setCity(shi);
                    y.setCommunitys(text);
                    y.setType(type);
                    r.add(y);

                }

            }

            System.out.println();

        }
        }
        return r;
    }


}
