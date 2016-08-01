package common.importandexport;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.Region;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CommonCell {
	//  /**  
	//   * 把一个excel中的cellstyletable复制到另一个excel，这里会报错，不能用这种方法，不明白呀？？？？？  
	//   * @param fromBook  
	//   * @param toBook  
	//   */  
	//  public static void copyBookCellStyle(HSSFWorkbook fromBook,HSSFWorkbook toBook){  
//	      for(short i=0;i<fromBook.getNumCellStyles();i++){  
//	          HSSFCellStyle fromStyle=fromBook.getCellStyleAt(i);  
//	          HSSFCellStyle toStyle=toBook.getCellStyleAt(i);  
//	          if(toStyle==null){  
//	              toStyle=toBook.createCellStyle();  
//	          }  
//	          copyCellStyle(fromStyle,toStyle);  
//	      }  
	//  }  
	    /** 
	     * 复制一个单元格样式到目的单元格样式 
	     * @param fromStyle 
	     * @param toStyle 
	     */  
	    public static void copyCellStyle(HSSFCellStyle fromStyle,  
	            HSSFCellStyle toStyle) {  
	        toStyle.setAlignment(fromStyle.getAlignment());  
	        //边框和边框颜色  
	        toStyle.setBorderBottom(fromStyle.getBorderBottom());  
	        toStyle.setBorderLeft(fromStyle.getBorderLeft());  
	        toStyle.setBorderRight(fromStyle.getBorderRight());  
	        toStyle.setBorderTop(fromStyle.getBorderTop());  
	        toStyle.setTopBorderColor(fromStyle.getTopBorderColor());  
	        toStyle.setBottomBorderColor(fromStyle.getBottomBorderColor());  
	        toStyle.setRightBorderColor(fromStyle.getRightBorderColor());  
	        toStyle.setLeftBorderColor(fromStyle.getLeftBorderColor());  
	          
	        //背景和前景  
	        toStyle.setFillBackgroundColor(fromStyle.getFillBackgroundColor());  
	        toStyle.setFillForegroundColor(fromStyle.getFillForegroundColor());  
	          
	        toStyle.setDataFormat(fromStyle.getDataFormat());  
	        toStyle.setFillPattern(fromStyle.getFillPattern());  
//	      toStyle.setFont(fromStyle.getFont(null));  
	        toStyle.setHidden(fromStyle.getHidden());  
	        toStyle.setIndention(fromStyle.getIndention());//首行缩进  
	        toStyle.setLocked(fromStyle.getLocked());  
	        toStyle.setRotation(fromStyle.getRotation());//旋转  
	        toStyle.setVerticalAlignment(fromStyle.getVerticalAlignment());  
	        toStyle.setWrapText(fromStyle.getWrapText());  
	          
	    }  
	    /** 
	     * Sheet复制 
	     * @param fromSheet 
	     * @param toSheet 
	     * @param copyValueFlag 
	     */  
	    public static void copySheet(HSSFWorkbook wb,HSSFSheet fromSheet, HSSFSheet toSheet,  
	            boolean copyValueFlag) {  
	        //合并区域处理  
	        mergerRegion(fromSheet, toSheet);  
	        for (Iterator rowIt = fromSheet.rowIterator(); rowIt.hasNext();) {  
	            HSSFRow tmpRow = (HSSFRow) rowIt.next();  
	            HSSFRow newRow = toSheet.createRow(tmpRow.getRowNum());  
	            //行复制  
	            copyRow(wb,tmpRow,newRow,copyValueFlag);  
	        }  
	    }  
	    /** 
	     * 行复制功能 
	     * @param fromRow 
	     * @param toRow 
	     */  
	    public static void copyRow(HSSFWorkbook wb,HSSFRow fromRow,HSSFRow toRow,boolean copyValueFlag){  
	        for (Iterator cellIt = fromRow.cellIterator(); cellIt.hasNext();) {  
	            HSSFCell tmpCell = (HSSFCell) cellIt.next();  
	            HSSFCell newCell = toRow.createCell(tmpCell.getCellNum());  
	            copyCell(wb,tmpCell, newCell, copyValueFlag);  
	        }  
	    }  
	    /** 
	    * 复制原有sheet的合并单元格到新创建的sheet 
	    *  
	    * @param sheetCreat 新创建sheet 
	    * @param sheet      原有的sheet 
	    */  
	    public static void mergerRegion(HSSFSheet fromSheet, HSSFSheet toSheet) {  
	       int sheetMergerCount = fromSheet.getNumMergedRegions();  
	       for (int i = 0; i < sheetMergerCount; i++) {  
	        Region mergedRegionAt = fromSheet.getMergedRegionAt(i);  
	        toSheet.addMergedRegion(mergedRegionAt);  
	       }  
	    }  
	    /** 
	     * 复制单元格 
	     *  
	     * @param srcCell 
	     * @param distCell 
	     * @param copyValueFlag 
	     *            true则连同cell的内容一起复制 
	     */  
	    public static void copyCell(HSSFWorkbook wb,HSSFCell srcCell, HSSFCell distCell,  
	            boolean copyValueFlag) {  
	        HSSFCellStyle newstyle=wb.createCellStyle();  
	        copyCellStyle(srcCell.getCellStyle(), newstyle);  
	       // distCell.setEncoding(srcCell);  
	        //样式  
	        distCell.setCellStyle(newstyle);  
	        //评论  
	        if (srcCell.getCellComment() != null) {  
	            distCell.setCellComment(srcCell.getCellComment());  
	        }  
	        // 不同数据类型处理  
	        int srcCellType = srcCell.getCellType();  
	        distCell.setCellType(srcCellType);  
	        if (copyValueFlag) {  
	            if (srcCellType == HSSFCell.CELL_TYPE_NUMERIC) {  
	                if (HSSFDateUtil.isCellDateFormatted(srcCell)) {  
	                    distCell.setCellValue(srcCell.getDateCellValue());  
	                } else {  
	                    distCell.setCellValue(srcCell.getNumericCellValue());  
	                }  
	            } else if (srcCellType == HSSFCell.CELL_TYPE_STRING) {  
	                distCell.setCellValue(srcCell.getRichStringCellValue());  
	            } else if (srcCellType == HSSFCell.CELL_TYPE_BLANK) {  
	                // nothing21  
	            } else if (srcCellType == HSSFCell.CELL_TYPE_BOOLEAN) {  
	                distCell.setCellValue(srcCell.getBooleanCellValue());  
	            } else if (srcCellType == HSSFCell.CELL_TYPE_ERROR) {  
	                distCell.setCellErrorValue(srcCell.getErrorCellValue());  
	            } else if (srcCellType == HSSFCell.CELL_TYPE_FORMULA) {  
	                distCell.setCellFormula(srcCell.getCellFormula());  
	            } else { // nothing29  
	            }  
	        }  
	    }  
	
	/*
	 * excel 2007
	 */
	    /** 
	     * excel 2007 行复制功能 
	     * @param fromRow 
	     * @param toRow 
	     */  
	    public static void copyRow(XSSFWorkbook wb,XSSFRow fromRow,XSSFRow toRow,boolean copyValueFlag){  
	        for (Iterator cellIt = fromRow.cellIterator(); cellIt.hasNext();) {  
	            XSSFCell tmpCell = (XSSFCell) cellIt.next();  
	            XSSFCell newCell = toRow.createCell(tmpCell.getColumnIndex());  
	            copyCell(wb,tmpCell, newCell, copyValueFlag);  
	        }  
	    }  
	    
	    
	    /** 
	     * excel 2007 复制单元格 
	     *  
	     * @param srcCell 
	     * @param distCell 
	     * @param copyValueFlag 
	     *            true则连同cell的内容一起复制 
	     */  
	    public static void copyCell(XSSFWorkbook wb,XSSFCell srcCell, XSSFCell distCell,  
	            boolean copyValueFlag) {  
	        XSSFCellStyle newstyle=wb.createCellStyle();  
	        copyCellStyle(srcCell.getCellStyle(), newstyle);  
	        //样式  
	        distCell.setCellStyle(newstyle);  
	        //评论  
	        if (srcCell.getCellComment() != null) {  
	            distCell.setCellComment(srcCell.getCellComment());  
	        }  
	        // 不同数据类型处理  
	        int srcCellType = srcCell.getCellType();  
	        distCell.setCellType(srcCellType);  
	        if (copyValueFlag) {  
	            if (srcCellType == HSSFCell.CELL_TYPE_NUMERIC) {  
	                if (HSSFDateUtil.isCellDateFormatted(srcCell)) {  
	                    distCell.setCellValue(srcCell.getDateCellValue());  
	                } else {  
	                    distCell.setCellValue(srcCell.getNumericCellValue());  
	                }  
	            } else if (srcCellType == HSSFCell.CELL_TYPE_STRING) {  
	                distCell.setCellValue(srcCell.getRichStringCellValue());  
	            } else if (srcCellType == HSSFCell.CELL_TYPE_BLANK) {  
	                // nothing21  
	            } else if (srcCellType == HSSFCell.CELL_TYPE_BOOLEAN) {  
	                distCell.setCellValue(srcCell.getBooleanCellValue());  
	            } else if (srcCellType == HSSFCell.CELL_TYPE_ERROR) {  
	                distCell.setCellErrorValue(srcCell.getErrorCellValue());  
	            } else if (srcCellType == HSSFCell.CELL_TYPE_FORMULA) {  
	                distCell.setCellFormula(srcCell.getCellFormula());  
	            } else { // nothing29  
	            }  
	        }  
	    }
	    
	    /** 
	     * excel 2007 复制一个单元格样式到目的单元格样式 
	     * @param fromStyle 
	     * @param toStyle 
	     */  
	    public static void copyCellStyle(XSSFCellStyle fromStyle,  
	            XSSFCellStyle toStyle) {  
	        toStyle.setAlignment(fromStyle.getAlignment());  
	        //边框和边框颜色  
	        toStyle.setBorderBottom(fromStyle.getBorderBottom());  
	        toStyle.setBorderLeft(fromStyle.getBorderLeft());  
	        toStyle.setBorderRight(fromStyle.getBorderRight());  
	        toStyle.setBorderTop(fromStyle.getBorderTop());  
	        toStyle.setTopBorderColor(fromStyle.getTopBorderColor());  
	        toStyle.setBottomBorderColor(fromStyle.getBottomBorderColor());  
	        toStyle.setRightBorderColor(fromStyle.getRightBorderColor());  
	        toStyle.setLeftBorderColor(fromStyle.getLeftBorderColor());  
	          
	        //背景和前景  
	        toStyle.setFillBackgroundColor(fromStyle.getFillBackgroundColor());  
	        toStyle.setFillForegroundColor(fromStyle.getFillForegroundColor());  
	          
	        toStyle.setDataFormat(fromStyle.getDataFormat());  
	        toStyle.setFillPattern(fromStyle.getFillPattern());  
//	      toStyle.setFont(fromStyle.getFont(null));  
	        toStyle.setHidden(fromStyle.getHidden());  
	        toStyle.setIndention(fromStyle.getIndention());//首行缩进  
	        toStyle.setLocked(fromStyle.getLocked());  
	        toStyle.setRotation(fromStyle.getRotation());//旋转  
	        toStyle.setVerticalAlignment(fromStyle.getVerticalAlignment());  
	        toStyle.setWrapText(fromStyle.getWrapText());  
	          
	    }  
	    
	    // 封装要显示的数据 如 map 和 list
	    
	    public String transferXSSFCellType(XSSFCell cell){
		       String result = new String();   
		       switch (cell.getCellType()) {   
		       case XSSFCell.CELL_TYPE_NUMERIC:// 数字类型   
		           if (HSSFDateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式   
		               SimpleDateFormat sdf = null;   
		               
		               if (cell.getCellStyle().getDataFormat() == HSSFDataFormat   
		                       .getBuiltinFormat("h:mm")) {   
		                   sdf = new SimpleDateFormat("HH:mm");   
		                } else {// 日期   
		                    sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");   
		                }   
		                Date date = cell.getDateCellValue();   
		                result = sdf.format(date);   
		            } else if (cell.getCellStyle().getDataFormat() == 58) {   
		                // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)   
		                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");   
		                double value = cell.getNumericCellValue();   
		                Date date = org.apache.poi.ss.usermodel.DateUtil   
		                        .getJavaDate(value);   
		                result = sdf.format(date);   
		            } else {   
		                double value = cell.getNumericCellValue();   
		                XSSFCellStyle style = cell.getCellStyle();   
		                DecimalFormat format = new DecimalFormat();   
		                String temp = style.getDataFormatString();   
		                // 单元格设置成常规   
		                if (temp.equals("General")) {   
		                    format.applyPattern("#");   
		                }   
		                result = format.format(value);   
		            }   
		            break;   
		        case XSSFCell.CELL_TYPE_STRING:// String类型   
		            result = cell.getRichStringCellValue().toString();   
		            break;   
		        case XSSFCell.CELL_TYPE_BLANK:   
		            result = "";   
		        default:   
		            result = "";   
		            break;   
		        }   
		        return result;   
	    }
	    
		/**
		 * 判断指定的单元格是否是合并单元格
		 * 
		 * @param sheet
		 * @param row
		 *            行下标
		 * @param column
		 *            列下标
		 * @return
		 */
		public boolean isMergedRegion(Sheet sheet, int row, int column) {
			int sheetMergeCount = sheet.getNumMergedRegions();
			for (int i = 0; i < sheetMergeCount; i++) {
				CellRangeAddress range = sheet.getMergedRegion(i);
				int firstColumn = range.getFirstColumn();
				int lastColumn = range.getLastColumn();
				int firstRow = range.getFirstRow();
				int lastRow = range.getLastRow();
				if (row >= firstRow && row <= lastRow) {
					if (column >= firstColumn && column <= lastColumn) {
						return true;
					}
				}
			}
			return false;
		}
		
		
		/**
		 * 获取合并单元格的值
		 * 
		 * @param sheet
		 * @param row
		 * @param column
		 * @return
		 */
		public String getMergedRegionValue(Sheet sheet, int row, int column) {
			int sheetMergeCount = sheet.getNumMergedRegions();

			for (int i = 0; i < sheetMergeCount; i++) {
				CellRangeAddress ca = sheet.getMergedRegion(i);
				int firstColumn = ca.getFirstColumn();
				int lastColumn = ca.getLastColumn();
				int firstRow = ca.getFirstRow();
				int lastRow = ca.getLastRow();
				if (row >= firstRow && row <= lastRow) {

					if (column >= firstColumn && column <= lastColumn) {
						Row fRow = sheet.getRow(firstRow);
						Cell fCell = fRow.getCell(firstColumn);
						return getCellValue(fCell);
					}
				}
			}

			return null;
		}
		
		/**
		 * 设置合并单元格的行
		 * 
		 * @param sheet
		 * @param row
		 * @param column
		 * @return
		 */
		public void setMergedRegionRow(XSSFSheet sheet, int row, int column,String provalue) {
			int sheetMergeCount = sheet.getNumMergedRegions();

			for (int i = 0; i < sheetMergeCount; i++) {
				CellRangeAddress ca = sheet.getMergedRegion(i);
				int firstColumn = ca.getFirstColumn();
				int lastColumn = ca.getLastColumn();
				int firstRow = ca.getFirstRow();
				int lastRow = ca.getLastRow();
				if (row-1 >= firstRow && row-1 <= lastRow) {

					if (column >= firstColumn && column <= lastColumn) {
						// 跨第1行第1个到第2行第1个单元格的操作为 
						//sheet.addMergedRegion(new Region(0,(short)0,1,(short)0)); 
						XSSFRow currentFRow = sheet.getRow(firstRow);
						XSSFCell currentCell = currentFRow.getCell(column);
						String cellValue =  getCellValue(currentCell);
						if(cellValue.equals(provalue) == true){
							CellRangeAddress region = new CellRangeAddress(firstRow, row, firstColumn, lastColumn);
							sheet.addMergedRegion(region);
							sheet.removeMergedRegion(i);
						}

					}
				}
			}
		}
		
		/**
		 * 获取单元格的值
		 * 
		 * @param cell
		 * @return
		 */
		public String getCellValue(Cell cell) {

			if (cell == null)
				return "";

			if (cell.getCellType() == Cell.CELL_TYPE_STRING) {

				return cell.getStringCellValue();

			} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {

				return String.valueOf(cell.getBooleanCellValue());

			} else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {

				return cell.getCellFormula();

			} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {

				return String.valueOf(cell.getNumericCellValue());

			}
			return "";
		}
		
		/**
		 * 获取跨列的列数
		 * @param yujie
		 * @return
		 */
		public int  getCommonExport(Sheet sheet ,int row , int column){
			int sheetMergeCount = sheet.getNumMergedRegions(); 
			int lastCol = 0;
	        for(int i = 0 ; i < sheetMergeCount ; i++){  
	            CellRangeAddress ca = sheet.getMergedRegion(i);  
	            int firstColumn = ca.getFirstColumn();  
	            int lastColumn = ca.getLastColumn();  
	            int firstRow = ca.getFirstRow();  
	            int lastRow = ca.getLastRow();  
	            if(row >= firstRow && row <= lastRow){  
	                if(column >= firstColumn && column <= lastColumn){  
	                    Row fRow = sheet.getRow(firstRow);  
	                    Cell fCell = fRow.getCell(firstColumn);  
	                       lastCol =  lastColumn;
	                       break;
	                }  
	            }else{
	            	continue;
	            }  
	        }
	        
			return lastCol;
		}
		
		  /** 
	     * Converts an Excel column name like "C" to a zero-based index. 
	     *  
	     * @param name 
	     * @return Index corresponding to the specified name 
	     */  
	    public int nameToColumn(String name) {  
	        int column = -1;  
	        for (int i = 0; i < name.length(); ++i) {  
	            int c = name.charAt(i);  
	            column = (column + 1) * 26 + c - 'A';  
	        }  
	        return column;  
	    }
	    
		// 转换一下方法名
		public static String convertToMethodName(String attribute,Class objClass,boolean isSet)
	    {
			//  attribute = attribute.toLowerCase();
			String REGEX = "[a-zA-Z]";
	        /** 通过正则表达式来匹配第一个字符 **/
	        Pattern p = Pattern.compile(REGEX);
	        Matcher m = p.matcher(attribute);
	        StringBuilder sb = new StringBuilder();
	        /** 如果是set方法名称 **/
	        if(isSet)
	        {
	            sb.append("set");
	        }else{
	        /** get方法名称 **/
	            try {
	                Field attributeField = objClass.getDeclaredField(attribute);
	                /** 如果类型为boolean **/
	                if(attributeField.getType() == boolean.class||attributeField.getType() == Boolean.class)
	                {
	                    sb.append("is");
	                }else
	                {
	                    sb.append("get");
	                }
	            } catch (SecurityException e) {
	                e.printStackTrace();
	            } catch (NoSuchFieldException e) {
	                e.printStackTrace();
	            }
	        }
	        /** 针对以下划线开头的属性 **/
	        if(attribute.charAt(0)!='_' && m.find())
	        {
	            sb.append(m.replaceFirst(m.group().toUpperCase()));
	        }else{
	            sb.append(attribute);
	        }
	        return sb.toString();
	    }
		/**
		 * 获得get方法的值
		 * @param obj
		 * @param attribute
		 * @return
		 */
		  public static Object getAttrributeValue(Object obj,String methodName)
		    {
		        //String methodName = convertToMethodName(attribute, obj.getClass(), false);
		        Object value = null;
		        try {
		            /** 由于get方法没有参数且唯一，所以直接通过方法名称锁定方法 **/
		            Method methods = obj.getClass().getDeclaredMethod(methodName);
		            if(methods != null)
		            {
		                value = methods.invoke(obj);
		            }
		        } catch (SecurityException e) {
		            e.printStackTrace();
		        } catch (NoSuchMethodException e) {
		            e.printStackTrace();
		        } catch (IllegalArgumentException e) {
		            e.printStackTrace();
		        } catch (IllegalAccessException e) {
		            e.printStackTrace();
		        } catch (InvocationTargetException e) {
		            e.printStackTrace();
		        }
		        return value;
		    }
	    
	   
}
