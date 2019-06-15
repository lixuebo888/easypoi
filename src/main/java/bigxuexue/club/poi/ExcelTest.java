package bigxuexue.club.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;

public class ExcelTest {

	@Test
	public void fe_map() throws Exception {
		TemplateExportParams params = new TemplateExportParams("doc/专项支出用款申请书2.xls");
		Map<String, Object> map = new HashMap<String, Object>();
		map.put("date", "201q411-12-25");
		map.put("money", 2000000.00);
		map.put("upperMoney", "贰佰万");
		map.put("company", "执笔潜行科技有限公司");
		map.put("bureau", "财政局");
		map.put("person", "JueYue");
		map.put("phone", "1879740****");
		List<Map<String, String>> listMap = new ArrayList<Map<String, String>>();
		for (int i = 0; i < 4; i++) {
			Map<String, String> lm = new HashMap<String, String>();
			lm.put("id", i + 1 + "");
			lm.put("zijin", i * 10000 + "");
			lm.put("bianma", "A001");
			lm.put("mingcheng", "设计");
			lm.put("xiangmumingcheng", "EasyPoi " + i + "期");
			lm.put("quancheng", "开源项目");
			lm.put("sqje", i * 10000 + "");
			lm.put("hdje", i * 10000 + "");
			lm.put("name", "李四" + i + "");
			lm.put("age", "age" + i + "");
			lm.put("add", "add" + i + "");

			listMap.add(lm);
		}
		map.put("maplist", listMap);

		Workbook workbook = ExcelExportUtil.exportExcel(params, map);
		File savefile = new File("D:/excel/");
		if (!savefile.exists()) {
			savefile.mkdirs();
		}
		FileOutputStream fos = new FileOutputStream("excelExport/专项支出用款申请书_map.xls");
		workbook.write(fos);
		fos.close();
	}
}
