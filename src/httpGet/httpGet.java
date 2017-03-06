package httpGet;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.URL;
import java.net.URLConnection;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import jxl.Cell;
import jxl.LabelCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import org.json.*; 
import org.apache.poi.ss.usermodel.Row; 


public class httpGet {
	public static String sendGet(String url){
		String result = "";
		 BufferedReader in = null;
		try {
			URL realUrl = new URL(url);
			URLConnection connection = realUrl.openConnection();
			connection.setRequestProperty("accept", "*/*");
            connection.setRequestProperty("connection", "Keep-Alive");
            connection.setRequestProperty("user-agent",
                    "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1;SV1)");
            
            connection.connect();
            in = new BufferedReader(new InputStreamReader(
                    connection.getInputStream()));
            String line;
            while ((line = in.readLine()) != null) {
                result += line;
            }
		} catch (Exception e) {
			// TODO: handle exception
			System.out.println("发送GET请求出现异常！" + e);
            e.printStackTrace();
		}finally {
            try {
                if (in != null) {
                    in.close();
                }
            } catch (Exception e2) {
                e2.printStackTrace();
            }
        }
		
		return result;
	}
	
	public static List<Double> getLngAndLat(String s) throws JSONException{
		System.out.println("json+"+s);
        List<Double> l1 = new ArrayList<Double>();
		JSONObject a = new JSONObject(s); 
		if(a.getInt("status")!=0){
			l1.add((double) 0);
			l1.add((double) 0);
			return l1;
		}
        String s1 = a.getString("result");
        JSONObject a1 = new JSONObject(s1);
        String s2 = a1.getString("location");
        JSONObject a2 = new JSONObject(s2);
        System.out.println(a2.get("lng"));
        

        
        l1.add( (Double) a2.get("lng"));
        l1.add((Double) a2.get("lat"));
        return l1;
	}

	 public static void main(String[] args) throws JSONException {
	        //发送 GET 请求
	        String s= new String();/*httpGet.sendGet("http://api.map.baidu.com/geocoder/v2/?output=json&address=渭南市渭清路南段&ak=ue9Q5ytuQZyq3t4HQip5VYY02qI9vah7");*/
	        String path1 = new String("http://api.map.baidu.com/geocoder/v2/?");
	        String path3 = new String("output=json&address=");
	        String path2 = new String("&ak=ue9Q5ytuQZyq3t4HQip5VYY02qI9vah7");
	        List<Double> l1 = new ArrayList<>();/*getLngAndLat(s);*/
/*	        System.out.println(l1.toString());*/
	        
	        jxl.Workbook readwb = null;
	        jxl.Workbook readwb1 = null;
	        
	        try {
	        	
	        	WorkbookSettings settings=new WorkbookSettings();
	        	settings.setEncoding("GBK");
	        	InputStream inputStream = new FileInputStream("D://1206.xls");
	        	InputStream inputStream1 = new FileInputStream("D://20162.xls");
	        	readwb = Workbook.getWorkbook(inputStream);
	        	readwb1 = Workbook.getWorkbook(inputStream1,settings);
	        	
	        	Sheet readSheet = readwb.getSheet(0);
	        	Sheet readSheet1 = readwb1.getSheet(0);
	        	
	        	int rscolumns = readSheet.getColumns();
				System.out.println("列数是："+rscolumns);
				
				int rsrows = readSheet.getRows();
				System.out.println("行数是："+rsrows);
				
				List<String> l2016= new ArrayList<>();
				List<String> l20161 = new ArrayList<>();
				List<String> l20162 = new ArrayList<>();
				for(int i =0 ;i<readSheet1.getRows();i++){
					l2016.add(readSheet1.getCell(1, i).getContents());
					l20161.add(readSheet1.getCell(3, i).getContents());
					l20162.add(readSheet1.getCell(2, i).getContents());
				}
				System.out.println("sdsad "+l20161.get(0));
				
				//新建写入文件
				File filewrite = new File("D://12061.xls");
				filewrite.createNewFile();
				OutputStream os = new FileOutputStream(filewrite);
				
				WritableWorkbook wwb = Workbook.createWorkbook(os);
				WritableSheet ws = wwb.createSheet("Sheet1",0);
				int i=0;
				while(i<rsrows){
					String string = readSheet.getCell(0, i).getContents();
					int j = l2016.indexOf(string);
					if(j!=-1){
						Label label = new Label(0, i, readSheet.getCell(0, i).getContents());
						ws.addCell(label);
						Label label4 = new Label(1, i, readSheet.getCell(1,i).getContents());
						ws.addCell(label4);
						Label label5 = new Label(2, i, readSheet.getCell(2, i).getContents());
						ws.addCell(label5);
						Label label6 = new Label(3, i, l20162.get(j));
						ws.addCell(label6);
						Label label1 = new Label(4, i, l20161.get(j));
						ws.addCell(label1);
						s = l20161.get(j);
						s= URLEncoder.encode(s, "UTF-8");
						s= path1+path3+s+path2;
						l1 = getLngAndLat(httpGet.sendGet(s));
						Label label2 = new Label(5, i, l1.get(0).toString());
						ws.addCell(label2);
						Label label3 = new Label(6, i, l1.get(1).toString());
						ws.addCell(label3);
						
					}
					else{
						Label label = new Label(0, i, readSheet.getCell(0, i).getContents());
						ws.addCell(label);
						Label label4 = new Label(1, i, readSheet.getCell(1,i).getContents());
						ws.addCell(label4);
						Label label5 = new Label(2, i, readSheet.getCell(2, i).getContents());
						ws.addCell(label5);
					}
					i++;
				}
				wwb.write();
				wwb.close();
			} catch (Exception e) {
				// TODO: handle exception
				e.printStackTrace();
			}finally{
				readwb.close();
				readwb1.close();
			}
	  }
}
