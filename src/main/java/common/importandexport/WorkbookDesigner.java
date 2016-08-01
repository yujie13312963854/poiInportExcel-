package common.importandexport;

import java.math.BigDecimal;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkbookDesigner {

	private XSSFWorkbook xSSFWorkbook;
	
	private Map  dataSource = new HashMap();

	public XSSFWorkbook getxSSFWorkbook() {
		return xSSFWorkbook;
	}

	public void setxSSFWorkbook(XSSFWorkbook xSSFWorkbook) {
		this.xSSFWorkbook = xSSFWorkbook;
	}

	public Map getDataSource() {
		return dataSource;
	}

	public void setDataSource(String key ,Object value) {
		
		dataSource.put(key, value);
		 
	}
	
	public void process(boolean varbl,int sheetnumber){
		if(varbl == true){
			XSSFSheet fromsheet = xSSFWorkbook.getSheetAt(sheetnumber);
			int lastRowNum = fromsheet.getLastRowNum();
			CommonCell commonCell = new CommonCell();
			// 匹配 行合计
			Pattern pattern=null;  
			Matcher macher=null; 
		    pattern=Pattern.compile("^[A-Z0-9\\+]{1,}$");  
	       
		    // 匹配行合计
		    Pattern sumPattern=null;  
			Matcher sumMatch=null; 
			sumPattern=Pattern.compile("^&=&=(Sum|SUM|sum)\\([A-Z]{1}\\{r\\}:[A-Z]{1}\\{r\\}\\)$");  
			
			  // 匹配列合计  =SUM(M8:M8)
			  Pattern colPattern=null;  
			  Matcher colMatch=null; 
			  colPattern=Pattern.compile("^(Sum|SUM|sum)\\([A-Z]{1}[0-9]{1,}:[A-Z]{1}[0-9]{1,}\\)$");
			  
			  //  匹配数字
			  Pattern dpattern=null;  
				Matcher dmdacher=null; 
				dpattern=Pattern.compile("^[-]{0,1}[0-9]+([.][0-9]+){0,1}$");  
				
			  
			  // 动态插入listOfSize-1 行
			  for(int i=0;i<=lastRowNum; i++){
				// 需要插入的行数
					int  listOfSize = 0;
				// 是否要插入多行
				boolean isInsertRow = false;
				
				XSSFRow fromFRow = fromsheet.getRow(i);
				
				if(fromFRow != null){
					int cellnum = fromFRow.getLastCellNum();
					for(int j=0;j< cellnum;){
						// 遍历每一个cell
						String XSSFCellvalue = "";
						XSSFCell  xSSFCell =  fromFRow.getCell(j);
						// 判断是否是合并单元格
						boolean bl = commonCell.isMergedRegion(fromsheet, i, j);
						if(bl == true){
							XSSFCellvalue = commonCell.getMergedRegionValue(fromsheet, i, j);
							int spanColum = new CommonCell().getCommonExport(fromsheet, i, j);
							// 每一个cell 的跨度
							j =  spanColum +1;
						}else{
							XSSFCellvalue = commonCell.getCellValue(xSSFCell);
							// 判断是否要插入多行
							XSSFRow nextFRow = fromsheet.getRow(i+1);
							// 判断下一行与当前行的值是否一致，判断上一行与当前行的值是否一致 若上一行与当前行的值不一致，当前行与下一行的值不一致 则漂移listOfSize行
							XSSFCell  nextCell = null;
							String nextCellvalue = "";
							if(nextFRow != null ){
							    nextCell = nextFRow.getCell(j);
							    if(commonCell.isMergedRegion(fromsheet, i+1, j)){
							    	 nextCellvalue = commonCell.getMergedRegionValue(fromsheet, i+1, j);
							    }else{
							    	 nextCellvalue = commonCell.getCellValue(nextCell);
							    }
							}
							
							XSSFCell  preCell = null;
							String preCellvalue =  "";
							XSSFRow preFRow = fromsheet.getRow(i-1);
							if(preFRow != null){
							   preCell = preFRow.getCell(j);
							   if(commonCell.isMergedRegion(fromsheet, i-1, j)){
								   preCellvalue = commonCell.getMergedRegionValue(fromsheet, i-1, j);
							    }else{
							    	preCellvalue = commonCell.getCellValue(preCell);
							    }
							}
							
							if(nextCell != null ){
								if(nextCellvalue.equals(XSSFCellvalue) == false && preCellvalue.equals(XSSFCellvalue) == false ){
									isInsertRow = true;
								}
							}else{
								if(preCellvalue.equals(XSSFCellvalue) == false ){
									isInsertRow = true;
								}
								
							}
							
							j++;
						}
						//  
						if(XSSFCellvalue.contains("&=") == true &&  XSSFCellvalue.contains("&=&=") == false){
							int startIndex =  XSSFCellvalue.indexOf("=");
							int endIndex =  XSSFCellvalue.indexOf(".");
							String varkey = XSSFCellvalue.substring(startIndex+1,endIndex);
							String key = varkey.trim();
							//获得参数的对象
							Object object = dataSource.get(key);   
							startIndex =  XSSFCellvalue.indexOf(".");
							if(object instanceof List){
								// list 现在目前不支持跨行
								List currentList = ((List) object);
								int currentSize = currentList.size();
								if(listOfSize < currentSize && isInsertRow == true){
									// 多个List 以size 最大的那个进行插入多行
									listOfSize = currentSize;

									
								}
							}
						}
						
					}
					//
					if(listOfSize != 0  && listOfSize != 1){
						// 动态插入的行数-- listOfSize
						// i+1 代表从第i+1行开始移动，fromsheet.getLastRowNum() 移动截至的行数， listOfSize 代表插入几行
						fromsheet.shiftRows(i, fromsheet.getLastRowNum(), listOfSize-1,true,false);  
						for(int m=0;m<listOfSize-1;m++){
							//  复制每一行的内容到新行
							XSSFRow sourceRow = fromsheet.createRow(i+m);
							// true 代表内容也一块插入
							commonCell.copyRow(xSSFWorkbook, fromFRow, sourceRow, true);
						   //总行数扩大listOfSize 行
					    }
						 lastRowNum = lastRowNum +listOfSize-1;
				    }
				}
				
			  }
			  
			//  每个list的索引
			int indexList = 0;  
			for(int i=0;i<=lastRowNum; i++){
				// 判断已经迭代到的List 的索引
				
				boolean isScrollIndex =false;
				XSSFRow fromFRow = fromsheet.getRow(i);
			    if(fromFRow != null){
					int cellnum = fromFRow.getLastCellNum();
					//  标识是否是另外一个list
					int isblank = 0;
					
					for(int j=0;j< cellnum;){
						// 遍历每一个cell
						String XSSFCellvalue = "";
						XSSFCell  xSSFCell =  fromFRow.getCell(j);
						// 判断是否是合并单元格
						boolean bl = commonCell.isMergedRegion(fromsheet, i, j);
						int currentJ = 0;
						if(bl == true){
							XSSFCellvalue = commonCell.getMergedRegionValue(fromsheet, i, j);
							int spanColum = new CommonCell().getCommonExport(fromsheet, i, j);
							// 每一个cell 的跨度
							currentJ = j;
							j =  spanColum +1;
							
							
							
						}else{
							XSSFCellvalue = commonCell.getCellValue(xSSFCell);
							currentJ = j;
							j++;
						}
						macher=pattern.matcher(XSSFCellvalue); 
						sumMatch=sumPattern.matcher(XSSFCellvalue); 
						colMatch=colPattern.matcher(XSSFCellvalue); 
						if(XSSFCellvalue.contains("&=") == true &&  XSSFCellvalue.contains("&=&=") == false){
							 putXssfCell(  xSSFCell, XSSFCellvalue, commonCell, indexList,i,currentJ, fromsheet);
						}else if(macher.find()){
							String cellvalue =  macher.group(0).substring(0,XSSFCellvalue.length());
							String[] cellvals = cellvalue.split("\\+");
							BigDecimal bigDecimal = new BigDecimal("0");
							// A5+B5+C5
							 for(String var:cellvals){
								 // 首拼音字母 
								 String letter = var.substring(0, 1);
								 
								 String rownumStr = var.substring(1, var.length());
								// 代表多少行
								 int rowNum = new Integer(rownumStr).intValue() ;
								// 代表多少列
								int columnum =  commonCell.nameToColumn(letter);
								
								XSSFRow currentRow = fromsheet.getRow(rowNum-1);
								XSSFCell currentCell =  currentRow.getCell(columnum);
							    String currentCellValue = commonCell.getCellValue(currentCell);
							   // currentCellValue = getXssfCellValue(currentCellValue,commonCell,indexList);
							  
							    dmdacher=dpattern.matcher(currentCellValue); 
							    if("".equals(currentCellValue) == false && currentCellValue != null && dmdacher.find()){
							    	bigDecimal = bigDecimal.add(new BigDecimal(currentCellValue));
							    	 
							    }
							   
							    
						       }
							 // 当前的  cell
							 xSSFCell.setCellValue(bigDecimal.toString());
						 }// 求合计  &=&=Sum(G{r}:I{r})  
						else if(XSSFCellvalue.contains("&=&=") == true && sumMatch.find() == true){
							String cellvalue =  sumMatch.group(0);
							int bracket = cellvalue.indexOf("{");
							String startLetter = cellvalue.substring(8, bracket);
							int colon = cellvalue.indexOf(":");
							int secondBracket = cellvalue.indexOf("{", colon);
							String endLetter = cellvalue.substring(colon+1, secondBracket);
							// 合计的开始列
							int startColum =  commonCell.nameToColumn(startLetter);
							// 合计的结束列
							int endColum =  commonCell.nameToColumn(endLetter);
							// 对开始列到结束列的值进行累加求出合计
							BigDecimal bigDecimal = new BigDecimal("0");
							// 判断是否所有cell 都为空
							
							int isblankAll = startColum;
							for(int k=startColum;k<= endColum;k++){
								XSSFCell currentCell = fromFRow.getCell(k);
								String currCellValue = commonCell.getCellValue(currentCell);
								dmdacher=dpattern.matcher(currCellValue); 
								
								
								if("".equals(currCellValue) == false && currCellValue != null && dmdacher.find() ){
									bigDecimal = bigDecimal.add(new BigDecimal(currCellValue));
								}else{
									isblankAll++;
								}
								 
							}
							if(isblankAll-1 ==endColum ){
								// 所在的合计列全为空
								//给合计列赋值
								xSSFCell.setCellValue("");
							}else{
								//给合计列赋值
								xSSFCell.setCellValue(bigDecimal.toString());
							}
							
							
							
						}
						else if(colMatch.find()){
							// =SUM(I8:I8)  求列合计
							String cellvalue =  colMatch.group(0);
							int indexLetter = cellvalue.indexOf("(");
							// 标识第几列的索引
							indexLetter =  indexLetter+1;
							// 获得该列属于第几列的拼音字母
							String collumCount = cellvalue.substring(indexLetter, indexLetter+1);
							// 标识列属于第几列
							int collInt = commonCell.nameToColumn(collumCount);
							// 标识从第几行开始求和的索引
                            //int countIndex = cellvalue.indexOf(":");  countIndex = countIndex-1; 
                          
							// // 对开始行到结束行的值进行累加求出合计
							BigDecimal bigDecimal = new BigDecimal("0");
							// 冒号索引
							int colonLetter = cellvalue.indexOf(":");
							// 
							String numCount = cellvalue.substring(indexLetter+1, colonLetter);
							int num = new Integer(numCount).intValue();
							num = num -1 ;
							num =i-indexList ;
//							for(int k=num;k<=indexList+num;k++){
//								XSSFRow colFRow = fromsheet.getRow(k);
//								if(colFRow != null ){
//									XSSFCell cell = colFRow.getCell(collInt);
//									
//									String currCellValue = commonCell.getCellValue(cell);
//									dmdacher=dpattern.matcher(currCellValue); 
//									
//									if("".equals(currCellValue) == false && currCellValue != null  && dmdacher.find()){
//										bigDecimal =  bigDecimal.add(new BigDecimal(currCellValue));
//									}
//								}
//							}
							for(int k=num;k<=indexList+num;k++){
							XSSFRow colFRow = fromsheet.getRow(k);
							if(colFRow != null ){
								XSSFCell cell = colFRow.getCell(collInt);
								
								String currCellValue = commonCell.getCellValue(cell);
								dmdacher=dpattern.matcher(currCellValue); 
								
								if("".equals(currCellValue) == false && currCellValue != null  && dmdacher.find()){
									bigDecimal =  bigDecimal.add(new BigDecimal(currCellValue));
								}
							}
						   }
							
							xSSFCell.setCellValue(bigDecimal.toString());
						}else{
							isblank++;
							// 如果所有的cell 都为"" ,就索引重新计算
							if(isblank == cellnum){
								indexList =0;
							}
						}
						if(XSSFCellvalue.contains("&=") == true &&  XSSFCellvalue.contains("&=&=") == false && isScrollIndex ==false){
							
							
							
							int startIndex =  XSSFCellvalue.indexOf("=");
							int endIndex =  XSSFCellvalue.indexOf(".");
							String varkey = XSSFCellvalue.substring(startIndex+1,endIndex);
							String key = varkey.trim();
							//获得参数的对象
							Object object = dataSource.get(key);   
							startIndex =  XSSFCellvalue.indexOf(".");
							//如果参数的对象是map 则 该值为key 若为list 则值为对象的属性
							String varvalue = XSSFCellvalue.substring(startIndex+1, XSSFCellvalue.length());
							
							if(object instanceof List){
								// 用来标识该行的索引已经下移
								isScrollIndex =true;
							}
						}
						
				     }
					if(isScrollIndex == true){
						// list 的索引
						indexList++;
					}
					
					
				}else{
					//  下一个list 的索引
					indexList=0;
				}
		    }
		//  删除一些空行
			


				
			 for(int i=0;i<=lastRowNum; i++){
				 XSSFRow fromFRow = fromsheet.getRow(i);
				 if(fromFRow != null ){
					 int cellnum = fromFRow.getLastCellNum();
						// 遍历每一个cell
							String XSSFCellvalue = "";
							int isAllBlank = 0;
							for(int j=0;j<cellnum;j++){
								XSSFCell  xSSFCell =  fromFRow.getCell(j);
								
								if(xSSFCell != null ){
									dmdacher=dpattern.matcher(commonCell.getCellValue(xSSFCell)); 
									
									sumMatch=sumPattern.matcher(commonCell.getCellValue(xSSFCell)); 
									colMatch=colPattern.matcher(commonCell.getCellValue(xSSFCell)); 
									if(dmdacher.find()  ||  sumMatch.find() ||  colMatch.find() ){
										XSSFRichTextString rechTextString =  xSSFCell.getRichStringCellValue();
									  //  XSSFCellStyle orignCellStyle =   xSSFWorkbook.createCellStyle();
									    XSSFDataFormat format= xSSFWorkbook.createDataFormat();
										XSSFCellStyle orignCellStyle =   xSSFCell.getCellStyle();
									    orignCellStyle.setDataFormat(format.getFormat("0.00"));
										xSSFCell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
										// xSSFCell.setCellStyle(orignCellStyle);
										 xSSFCell.setCellValue(Double.parseDouble(rechTextString.toString()));
									} 
								}
							   
							    
								String currCellValue = commonCell.getCellValue(xSSFCell);
								if("".equals(currCellValue) == true || null ==currCellValue){
									isAllBlank++;
								}else if(currCellValue.contains("&=") == true){
									xSSFCell.setCellValue("");
								}
							}
							if(isAllBlank==cellnum){
								fromsheet.removeRow(fromFRow);
								// i+1 代表从第i+1行开始移动，fromsheet.getLastRowNum() 移动截至的行数， listOfSize 代表插入几行
								//fromsheet.shiftRows(i+1, fromsheet.getLastRowNum(), -1,true,false);  
					      }
				}
		  }
			//  删除一些&=字符串
			 
			 
			 
	    }
	 }
	
	// 给元素Cell 替换真正的值
	public void putXssfCell(XSSFCell  xSSFCell,String XSSFCellvalue,CommonCell commonCell,int indexList,int row ,int colum,XSSFSheet fromsheet ){
		  //  匹配数字
		  Pattern dpattern=null;  
		  Matcher dmdacher=null; 
		  dpattern=Pattern.compile("^[0-9]+([.][0-9]+){0,1}$");  
		  
		int startIndex =  XSSFCellvalue.indexOf("=");
		int endIndex =  XSSFCellvalue.indexOf(".");
		String varkey = XSSFCellvalue.substring(startIndex+1,endIndex);
		String key = varkey.trim();
		//获得参数的对象
		Object object = dataSource.get(key);   
		startIndex =  XSSFCellvalue.indexOf(".");
		//如果参数的对象是map 则 该值为key 若为list 则值为对象的属性
		String varvalue = XSSFCellvalue.substring(startIndex+1, XSSFCellvalue.length());
		
		if(object instanceof MapDataSource){
			// 
			Map map = ((MapDataSource) object).getMap();
			String value = (String)map.get(varvalue);
			// xSSFCell.setCellValue(value);
			
			if(value != null && "".equals(value.toString()) == false){
				xSSFCell.setCellValue(value);
			}
			commonCell.setMergedRegionRow(fromsheet, row, colum,value);
			
		}else if(object instanceof List){
			// list 现在目前不支持跨行
			List currentList = ((List) object);
			if(currentList.size() <= indexList){
				xSSFCell.setCellValue("");
				return ;
			}
			
			Object obj = currentList.get(indexList);
			if(obj != null){
				// 利用反射 得到对象方法的名  
				String methodName = commonCell.convertToMethodName(varvalue,obj.getClass(),false);
				Object attributeValue = commonCell.getAttrributeValue(obj, methodName);
				if(attributeValue != null && "".equals(attributeValue.toString()) == false){
					BigDecimal bigDecimal = new BigDecimal("0");
					try {
						int sign = attributeValue.toString().indexOf(".");
						if(sign != -1){
							dmdacher=dpattern.matcher(attributeValue.toString()); 	
							if(attributeValue.toString().length()  >= sign+5  &&  dmdacher.find()){
								String attrValue =   attributeValue.toString().substring(0,sign+5);
								bigDecimal = bigDecimal.add(new BigDecimal(attrValue) );
							}
						}
					} catch (Exception e) {
					   e.printStackTrace();
					}
					if(bigDecimal.toString().equals("0") == true){
						xSSFCell.setCellValue(attributeValue.toString());
					}else{
						xSSFCell.setCellValue(bigDecimal.toString());
					}
				}
				
			}
		 }
	}

	public static void main(String[] args) {
		  Pattern dpattern=null;  
		 Matcher dmdacher=null; 
		 dpattern=Pattern.compile("^[-]{0,1}[0-9]+([.][0-9]+){0,1}$");  
		 
		 dmdacher = dpattern.matcher("4.0982");
		 
		System.out.println(dmdacher.find()); 
	}
	
}
