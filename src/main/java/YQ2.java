/**
 * @program: ExeclUtil
 * @description:
 * @author: soulx
 * @create: 2022-07-25 13:37
 **/
public class YQ2 {

	private String	province;

	/**
	 * <pre>
	 * 天津市
	 * </pre>
	 */
	private String	city;

	// 0 高 1 中
	private int type ;


	private String communitys;
	private boolean b;
	private String fx;

	public String getFx() {
		return fx;
	}

	public void setFx(String fx) {
		this.fx = fx;
	}

	public boolean isB() {
		return b;
	}

	public void setB(boolean b) {
		this.b = b;
	}

	public String getProvince() {
		return province;
	}

	public void setProvince(String province) {
		this.province = province;
	}

	public String getCity() {
		return city;
	}

	public void setCity(String city) {
		this.city = city;
	}

	public int getType() {
		return type;
	}

	public void setType(int type) {
		this.type = type;
	}

	public String getCommunitys() {
		return communitys;
	}

	public void setCommunitys(String communitys) {
		this.communitys = communitys;
	}
}
