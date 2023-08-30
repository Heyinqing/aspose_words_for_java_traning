
/**
 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */
package com.rw.utile;

import com.aspose.words.Document;
import com.aspose.words.net.System.Data.DataRelation;
import com.aspose.words.net.System.Data.DataRow;
import com.aspose.words.net.System.Data.DataSet;
import com.aspose.words.net.System.Data.DataTable;

import java.util.Map;
import java.util.function.Function;

/**
* @desc: spring-mybatis-demo
* @name: MultiExport2Doc.java
* @author: tompai
* @email：liinux@qq.com
* @createTime: 2020年4月11日 下午4:17:41
* @history:
* @version: v1.0
*/

public class MultiExport2Doc {


	private ExcelUtils excelUtils;
	/**
	* @author: tompai
	* @createTime: 2020年4月11日 下午4:17:41
	* @history:
	* @param args void
	*/

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		// 可以是doc或docx
		String template = "./template/test123.docx";
		// 可以是doc或docx
		String destdoc = "./test123_new.docx";
		ExcelUtils excelUtils1 = new ExcelUtils();
		try {
			if (!excelUtils1.GetWordLicense()) {
				return;
			}
			Document doc = new Document(template);
			// 主要调用aspose.words的邮件合并接口MailMerge
			// 3.1 填充单个文本域
			String[] Flds = new String[] { "Title", "Name", "URL", "Note" }; // 文本域
			String Name = "中玮科技";
			String URL = "http://zgwe.com";
			String Note = "让科技改变生活";
			Object[] Vals = new Object[] { "zgwe@2020", Name, URL, Note }; // 值
			doc.getMailMerge().execute(Flds, Vals); // 调用接口

			// 3.2 填充单层循环的表格
			DataTable visitTb = new DataTable("Visit"); // 网站访问量表格
			visitTb.getColumns().add("Date"); // 0 增加三个列 日期
			visitTb.getColumns().add("IP"); // 1 IP访问数量
			visitTb.getColumns().add("PV"); // 2 页面浏览量
			// 向表格中填充数据
			for (int i = 1; i < 15; i++) {
				DataRow row = visitTb.newRow(); // 新增一行
				row.set(0, "2020年2月" + i + "日"); // 根据列顺序填入数值
				row.set(1, i * 300);
				row.set(2, i * 400);
				visitTb.getRows().add(row); // 加入此行数据
			}
			// 对于无数据的情况，增加一行空记录
			if (visitTb.getRows().getCount() == 0) {
				DataRow row = visitTb.newRow();
				visitTb.getRows().add(row);
			}
			doc.getMailMerge().executeWithRegions(visitTb); // 调用接口

			// 3.3 填充具有两层循环的表格
			// 需要定义两个数据表格，且两者之间通过某列关联起来
			DataTable userTb = new DataTable("User"); // 用户
			userTb.getColumns().add("Name"); // 用户名称
			userTb.getColumns().add("RegDate");
			DataTable infoTb = new DataTable("Info"); // 用户信息
			infoTb.getColumns().add("Name"); // 用户名称 通过此列和上个表User关联
			infoTb.getColumns().add("Date");
			infoTb.getColumns().add("Time");

			// 3.3.1 填充用户信息
			for (int i = 1; i < 4; i++) {
				DataRow row = userTb.newRow();
				row.set(0, "User" + i);
				row.set(1, "2020年3月" + i + "日");
				userTb.getRows().add(row);
			}
			// 3.3.2 填充详细信息
			for (int i = 1; i < 6; i++) {
				for (int j = 1; j < 5; j++) {
					DataRow row = infoTb.newRow();
					row.set(0, "User" + i);
					row.set(1, "2020年1月" + j + "日");
					row.set(2, j * 2 * i);
					infoTb.getRows().add(row);
				}
			}
			// 3.3.3 将 User 和 Info 关联起来
			DataSet userSet = new DataSet();
			userSet.getTables().add(userTb);
			userSet.getTables().add(infoTb);
			String[] contCols = { "Name" };
			String[] lstCols = { "Name" };
			userSet.getRelations().add(new DataRelation("UserInfo", userTb, infoTb, contCols, lstCols));
			doc.getMailMerge().executeWithRegions(userSet); // 调用接口
			// 第四步 保存新word文档
			doc.save(destdoc);

			System.out.println("完成");
		} catch (Exception e) {

			// TODO Auto-generated catch block
			e.printStackTrace();

		} // 定义文档接口
	}

}
