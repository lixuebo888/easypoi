package bigxuexue.club.poi;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;

import cn.afterturn.easypoi.word.WordExportUtil;

public class PoiUtil {

	private static SimpleDateFormat format = new SimpleDateFormat("yyyy年MM月dd");

	/**
	 * 简单导出没有图片和Excel
	 */
	@Test
	public void SimpleWordExport() {
		Map<String, Object> map = new HashMap<String, Object>();
		map.put("department", "Easypoi");
		map.put("person", "JueYue");
		map.put("time", format.format(new Date()));
		map.put("me", "JueYue");
		map.put("date", "2015-01-03");
		try {
			XWPFDocument doc = WordExportUtil.exportWord07("word/Simple.docx", map);
			FileOutputStream fos = new FileOutputStream("wordExport/simple.docx");
			doc.write(fos);
			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	
}
