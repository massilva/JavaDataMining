package data;

import java.util.ArrayList;

public class Data {
	@SuppressWarnings("rawtypes")
	public ArrayList<ArrayList> data;

	public Data() {
		data = new ArrayList<ArrayList>();
	}

	@SuppressWarnings("rawtypes")
	public ArrayList<ArrayList> getData() {
		return data;
	}

	@SuppressWarnings("rawtypes")
	public void setData(ArrayList<ArrayList> data) {
		this.data = data;
	}
	
}
