import java.util.List;

/**
 * <pre>
 *  Root
 * </pre>
 * @author toolscat.com
 * @verison $Id: Root v 0.1 2022-07-23 21:11:17
 */
public class YQ{
    private boolean b;
    private String fx;
    /**
     * <pre>
     *
     * </pre>
     */
    private String	type;

    /**
     * <pre>
     * 天津市
     * </pre>
     */
    private String	province;

    /**
     * <pre>
     * 天津市
     * </pre>
     */
    private String	city;

    /**
     * <pre>
     * 河西区
     * </pre>
     */
    private String	county;

    /**
     * <pre>
     * 天津市 天津市 河西区
     * </pre>
     */
    private String	area_name;

    public String getFx() {
        return fx;
    }

    public void setFx(String fx) {
        this.fx = fx;
    }

    /**
     * <pre>
     * communitys
     * </pre>
     */
    private List<String> communitys;

    public boolean getB() {
        return b;
    }

    public void setB(Boolean b) {
        this.b = b;
    }

    public String getType() {
        return this.type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getProvince() {
        return this.province;
    }

    public void setProvince(String province) {
        this.province = province;
    }

    public String getCity() {
        return this.city;
    }

    public void setCity(String city) {
        this.city = city;
    }

    public String getCounty() {
        return this.county;
    }

    public void setCounty(String county) {
        this.county = county;
    }

    public String getArea_name() {
        return this.area_name;
    }

    public void setArea_name(String area_name) {
        this.area_name = area_name;
    }

    public List<String> getCommunitys() {
        return this.communitys;
    }

    public void setCommunitys(List<String> communitys) {
        this.communitys = communitys;
    }

}


