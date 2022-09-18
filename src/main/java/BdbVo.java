import java.util.List;

/**
 * @program: ExeclUtil
 * @description:
 * @author: soulx
 * @create: 2022-09-08 17:17
 **/
public class BdbVo {
  private int count;
  private List<YqOb> data;
	private List<City> city;

	public int getCount() {
		return count;
	}

	public void setCount(int count) {
		this.count = count;
	}

	public List<YqOb> getData() {
		return data;
	}

	public void setData(List<YqOb> data) {
		this.data = data;
	}

	public List<City> getCity() {
		return city;
	}

	public void setCity(List<City> city) {
		this.city = city;
	}

	public static class City {
		private String cityname;
		private String fullname;
		public void setCityname(String cityname) {
			this.cityname = cityname;
		}
		public String getCityname() {
			return cityname;
		}

		public void setFullname(String fullname) {
			this.fullname = fullname;
		}
		public String getFullname() {
			return fullname;
		}
	}
}
