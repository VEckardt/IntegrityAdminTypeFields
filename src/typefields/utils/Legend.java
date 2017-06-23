package typefields.utils;

import java.util.LinkedHashMap;

@SuppressWarnings("serial")
public class Legend extends LinkedHashMap<String, String> {

  public void Add(String s1, String s2) {
	super.put(s1, s2);
  }

  public String getLegend() {

	String data = "<u>Legend:</u><table border=1>";

	for (Object key : this.keySet()) {
	  data = data + "<tr><td>" + key + ":</td><td>" + this.get(key) + "</td></tr>";
	}
	return data + ("<table><br>");

  }
}
