import java.util.List;

/**
 * @program: ExeclUtil
 * @description:
 * @author: soulx
 * @create: 2022-09-08 17:20
 **/
public class YqOb {
	private String province;
	private String cityname;
	private String citycode;
	private List<Qu> qu;
	private int count;

	public String getProvince() {
		return province;
	}

	public void setProvince(String province) {
		this.province = province;
	}

	public String getCityname() {
		return cityname;
	}

	public void setCityname(String cityname) {
		this.cityname = cityname;
	}

	public String getCitycode() {
		return citycode;
	}

	public void setCitycode(String citycode) {
		this.citycode = citycode;
	}

	public List<Qu> getQu() {
		return qu;
	}

	public void setQu(List<Qu> qu) {
		this.qu = qu;
	}

	public int getCount() {
		return count;
	}

	public void setCount(int count) {
		this.count = count;
	}

	public static class Qu {
		private String qu;
		private String map;
		public void setQu(String qu) {
			this.qu = qu;
		}
		public String getQu() {
			return qu;
		}

		public void setMap(String map) {
			this.map = map;
		}
		public String getMap() {
			return map;
		}
	}
}
