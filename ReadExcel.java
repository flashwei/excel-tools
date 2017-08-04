package com.excel.utils.excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import com.excel.utils.StringUtil;
import com.excel.utils.excel.annotion.ExcelField;

/**
 * 读取excel文件内容 
 * 仅支持xlsx格式
 * @author Vachel.Wang
 * @date 2016年6月16日 下午5:10:18  
 * @version V1.2
 */
public class ReadExcel {
	
	private XSSFWorkbook workbook = null;
	private XSSFSheet sheet;
    private XSSFRow row;
    private int sheetIndex = 0 ;
    private int rowIndex = 0 ;
    private int cellIndex = 0 ;
   
    
	public static ReadExcel me(InputStream s) throws IOException{
		return new ReadExcel(s);
	}
	
	public ReadExcel(InputStream s) throws IOException{
			workbook = new XSSFWorkbook(s);
			setSheetIndex(0);
	}
	
	public <T> List<T> readCoustomTemplet(Class<T> clazz) throws InstantiationException, IllegalAccessException {
	    List<T> list = new ArrayList<T>();
	    
	    XSSFRow titleRow = sheet.getRow(rowIndex-1);
	    ArrayList<String> rowArr = new ArrayList<String>();
	    int cells = titleRow.getPhysicalNumberOfCells();
	    for (int j = cellIndex; j < cells; j++) {
            rowArr.add(getStringCellValue(titleRow.getCell(j)));
        }
	    
	    List<Field> allFields = getMappedFiled(clazz, null);
        Map<String, Field> fieldsMap = new HashMap<String, Field>();//定义一个map用于存放列的序号和field.
        for (Field field : allFields) {
            if (field.isAnnotationPresent(ExcelField.class)) {
                ExcelField attr = field.getAnnotation(ExcelField.class);
                field.setAccessible(true);
                fieldsMap.put(attr.title(), field);
            }
        }
        
        //row = sheet.getRow(rowIndex);
        for (int i = rowIndex ; i < sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            T entity = null;
            for (int j = cellIndex; j < cells; j++) {
                String c = getStringCellValue(row.getCell(j));
                entity = (entity == null ? clazz.newInstance() : entity);// 如果不存在实例则新建.
                Field field = fieldsMap.get(rowArr.get(j));// 从map中得到对应列的field.
               
                if (field==null) {
                    continue;
                }
                
                Class<?> fieldType = field.getType();
                if (String.class == fieldType) {
                    field.set(entity, String.valueOf(c));
                } else if ((Integer.TYPE == fieldType) || (Integer.class == fieldType)) {
                    field.set(entity, Integer.parseInt(c));
                } else if ((Long.TYPE == fieldType) || (Long.class == fieldType)) {
                    field.set(entity, Long.valueOf(c));
                } else if ((Float.TYPE == fieldType) || (Float.class == fieldType)) {
                    field.set(entity, Float.valueOf(c));
                } else if ((Short.TYPE == fieldType) || (Short.class == fieldType)) {
                    field.set(entity, Short.valueOf(c));
                } else if ((Double.TYPE == fieldType) || (Double.class == fieldType)) {
                    field.set(entity, Double.valueOf(c));
                } else if (Character.TYPE == fieldType) {
                    if ((c != null) && (c.length() > 0)) {
                        field.set(entity, Character.valueOf(c.charAt(0)));
                    }
                }
            }
            if (entity != null) {
                list.add(entity);
            }
        }
	    
	    return list;
    }
	
	private List<Field> getMappedFiled(Class clazz, List<Field> fields) {
        if (fields == null) {
            fields = new ArrayList<Field>();
        }

        Field[] allFields = clazz.getDeclaredFields();// 得到所有定义字段
        // 得到所有field并存放到一个list中.
        for (Field field : allFields) {
            if (field.isAnnotationPresent(ExcelField.class)) {
                fields.add(field);
            }
        }
        if (clazz.getSuperclass() != null && !clazz.getSuperclass().equals(Object.class)) {
            getMappedFiled(clazz.getSuperclass(), fields);
        }

        return fields;
    }
	
	/**
	 * 读取自定义读取方式数据
	 * @return
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public List<ArrayList<String>> readCoustom() throws FileNotFoundException, IOException{
		List<ArrayList<String>> excelData = new ArrayList<ArrayList<String>>();
		for (int i = rowIndex ; i < sheet.getPhysicalNumberOfRows(); i++) {
			ArrayList<String> rowArr = new ArrayList<String>();
			row = sheet.getRow(i);
			for (int j = cellIndex; j < row.getPhysicalNumberOfCells(); j++) {
				rowArr.add(getStringCellValue(row.getCell(j)));
			}
			excelData.add(rowArr);
		}
		return excelData;
	}
	/**
	 * 以title列为标准读取自定义数据
	 * @param titleIndex
	 * @return
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public List<ArrayList<String>> readCoustomByTitle(int titleIndex) throws FileNotFoundException, IOException{
		XSSFRow titleRow = sheet.getRow(titleIndex);
		List<ArrayList<String>> excelData = new ArrayList<ArrayList<String>>();
		for (int i = rowIndex ; i < sheet.getPhysicalNumberOfRows(); i++) {
			ArrayList<String> rowArr = new ArrayList<String>();
			row = sheet.getRow(i);
			for (int j = cellIndex; j < titleRow.getPhysicalNumberOfCells(); j++) {
				rowArr.add(getStringCellValue(row.getCell(j)));
			}
			excelData.add(rowArr);
		}
		return excelData;
	}
	
	/**
	 * 读取自定义读取非空数据
	 * @return
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public List<ArrayList<String>> readCoustomExcludeNull() throws FileNotFoundException, IOException{
		List<ArrayList<String>> excelData = new ArrayList<ArrayList<String>>();
		for (int i = rowIndex ; i < sheet.getPhysicalNumberOfRows(); i++) {
			ArrayList<String> rowArr = new ArrayList<String>();
			row = sheet.getRow(i);
			for (int j = cellIndex; j < row.getPhysicalNumberOfCells(); j++) {
				String cellValue = getStringCellValue(row.getCell(j)); 
				if(!cellValue.equals(""))
					rowArr.add(cellValue);
			}
			excelData.add(rowArr);
		}
		return excelData;
	}
	
	/**
	 * 读取所有
	 * @return
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public List<ArrayList<String>> readAll() throws FileNotFoundException, IOException{
		List<ArrayList<String>> excelData = new ArrayList<ArrayList<String>>();
		for (int i = 0 ; i < sheet.getPhysicalNumberOfRows(); i++) {
			ArrayList<String> rowArr = new ArrayList<String>();
			row = sheet.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				rowArr.add(getStringCellValue(row.getCell(j)));
			}
			excelData.add(rowArr);
		}
		return excelData;
	}
	
	/**
	 * 以title行为标准来读其它所有取列值
	 * @return
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public List<ArrayList<String>> readAllByTitle(int titleIndex) throws FileNotFoundException, IOException{
		XSSFRow titleRow = sheet.getRow(titleIndex);
		List<ArrayList<String>> excelData = new ArrayList<ArrayList<String>>();
		for (int i = 0 ; i < sheet.getPhysicalNumberOfRows(); i++) {
			ArrayList<String> rowArr = new ArrayList<String>();
			row = sheet.getRow(i);
			if(row==null)
				continue;
			for (int j = 0; j < titleRow.getPhysicalNumberOfCells(); j++) {
				rowArr.add(getStringCellValue(row.getCell(j)));
			}
			excelData.add(rowArr);
		}
		return excelData;
	}

	/**
	 * 读取title行下的所有数据
	 * @param titleIndex
	 * @return
	 * @throws FileNotFoundException
	 * @throws IOException
     */
	public List<ArrayList<String>> readNextByTitle(int titleIndex) throws FileNotFoundException, IOException{
		XSSFRow titleRow = sheet.getRow(titleIndex);
		List<ArrayList<String>> excelData = new ArrayList<ArrayList<String>>();
		for (int i = titleIndex+1 ; i < sheet.getPhysicalNumberOfRows(); i++) {
			ArrayList<String> rowArr = new ArrayList<String>();
			row = sheet.getRow(i);
			if(row==null){
			    continue;
			}
			for (int j = 0; j < titleRow.getPhysicalNumberOfCells(); j++) {
				rowArr.add(getStringCellValue(row.getCell(j)));
			}
			excelData.add(rowArr);
		}
		return excelData;
	}
	
	/**
	 * 读取非空数据
	 * @return
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public List<ArrayList<String>> readAllExcludeNull() throws FileNotFoundException, IOException{
		List<ArrayList<String>> excelData = new ArrayList<ArrayList<String>>();
		for (int i = 0 ; i < sheet.getPhysicalNumberOfRows(); i++) {
			ArrayList<String> rowArr = new ArrayList<String>();
			row = sheet.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				String cellValue = getStringCellValue(row.getCell(j)); 
				if(!cellValue.equals(""))
					rowArr.add(cellValue);
			}
			excelData.add(rowArr);
		}
		return excelData;
	}
	
	/**
	 * 读取行的所有列的数据
	 * @param rowIndex
	 * @return
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public ArrayList<String> readRowData(int rowIndex) throws FileNotFoundException, IOException{
		ArrayList<String> rowDataArr = new ArrayList<String>();
		row = sheet.getRow(rowIndex);
		for (int i = 0 ; i < row.getPhysicalNumberOfCells(); i++) {
			String cellStr = getStringCellValue(row.getCell(i));
			rowDataArr.add(cellStr);
		}
		return rowDataArr;
	}
	
	/**
	 * 读取行的所有列的非空数据
	 * @param rowIndex
	 * @return
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public ArrayList<String> readRowExcludeNullData(int rowIndex) throws FileNotFoundException, IOException{
		ArrayList<String> rowDataArr = new ArrayList<String>();
		row = sheet.getRow(rowIndex);
		for (int i = 0 ; i < row.getPhysicalNumberOfCells(); i++) {
			String cellValue = getStringCellValue(row.getCell(i));
			if(!cellValue.equals(""))
				rowDataArr.add(cellValue);
		}
		return rowDataArr;
	}
	
	/**
	 * 读取读列的数据 
	 * 空为 ""
	 * @param rowIndex
	 * @return
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public String readCellData(int rowIndex , int cellIndex) throws FileNotFoundException, IOException{
		row = sheet.getRow(rowIndex);
		return getStringCellValue(row.getCell(cellIndex));
	}
	
	/**
	 * 获取所有图片 
	 * @return
	 */
	public List<XSSFPictureData> readPictures() {
		return workbook.getAllPictures();
	}
	
	/**
	 * 读取单元格内容，支持获取函数内容
	 * @param cell
	 * @return
	 */
	public static String getStringCellValue(Cell cell) {
        String strCell = "";
        if(cell==null) return strCell;
        switch (cell.getCellType()) {
	        case Cell.CELL_TYPE_STRING:	// get String data 
	        	strCell = cell.getRichStringCellValue().getString().trim();
	            break;
	        case Cell.CELL_TYPE_NUMERIC:	// get date or number data 
	            if (DateUtil.isCellDateFormatted(cell)) {
	                strCell = com.excel.utils.DateUtil.formatYYYYMMDDHHMMSS(cell.getDateCellValue());
	            } else {
	            	strCell = String.valueOf(cell.getNumericCellValue());
	            }
	            break;
	        case Cell.CELL_TYPE_BOOLEAN:	// get boolean data 
	        	strCell = String.valueOf(cell.getBooleanCellValue());
	            break;
	        case Cell.CELL_TYPE_FORMULA:	// get expression data 
	        	FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
	            evaluator.evaluateFormulaCell(cell);
	            CellValue cellValue = evaluator.evaluate(cell);
				if(StringUtil.isNotBlank(cellValue.getStringValue())){
					strCell = String.valueOf(cellValue.getStringValue()) ;
				}else{
					strCell = String.valueOf(cellValue.getNumberValue()) ;
				}
	            break;
	        default:
	        	strCell = "";
	    }
        return strCell;
    }
	
	/**
	 * 设置读取sheet下标，默认为0
	 * @param sheetIndex
	 * @return
	 */
	public ReadExcel setSheetIndex(int sheetIndex){
		this.sheetIndex = sheetIndex ;
		sheet = workbook.getSheetAt(sheetIndex);
		return this;
	}
	/**
	 * 设置读取sheet下标，默认为0
	 * @param rowIndex
	 * @return
	 */
	public ReadExcel setRowIndex(int rowIndex){
		this.rowIndex = rowIndex ;
		return this;
	}
    /**
	 * 设置读取cell下标，默认为0
	 * @param rowIndex
	 * @return
	 */
	public ReadExcel setCellIndex(int cellIndex){
		this.cellIndex = cellIndex ;
		return this;
	}
	/**
	 * 获取sheet总数
	 * @return
	 */
	public int getSheetCount(){
		return this.workbook.getNumberOfSheets();
	}
	/**
	 * 根据sheet下标获取sheet名称
	 * @param index
	 * @return
	 */
	public String getSheetNameByIndex(int index){
		return this.workbook.getSheetAt(index).getSheetName();
	}
	public String getCurrentSheetName(){
		return this.workbook.getSheetAt(sheetIndex).getSheetName();
	}
	/**
	 * 获取Excel2007图片
	 * @param sheetNum 当前sheet编号
	 * @param sheet 当前sheet对象
	 * @param workbook 工作簿对象
	 * @return Map key:图片单元格索引（0_1_1）String，value:图片流PictureData
	 */
	public static Map<String, PictureData> getSheetPictrues(int sheetNum,
															  XSSFSheet sheet, XSSFWorkbook workbook) {
		Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();

		for (POIXMLDocumentPart dr : sheet.getRelations()) {
			if (dr instanceof XSSFDrawing) {
				XSSFDrawing drawing = (XSSFDrawing) dr;
				List<XSSFShape> shapes = drawing.getShapes();
				for (XSSFShape shape : shapes) {
					XSSFPicture pic = (XSSFPicture) shape;
					XSSFClientAnchor anchor = pic.getPreferredSize();
					CTMarker ctMarker = anchor.getFrom();
					String picIndex = String.valueOf(sheetNum) + "_"
							+ ctMarker.getRow() + "_" + ctMarker.getCol();
					sheetIndexPicMap.put(picIndex, pic.getPictureData());
				}
			}
		}

		return sheetIndexPicMap;
	}

	public static void printImg(List<Map<String, XSSFPictureData>> sheetList) throws IOException {

		for (Map<String, XSSFPictureData> map : sheetList) {
			Object key[] = map.keySet().toArray();
			for (int i = 0; i < map.size(); i++) {
				// 获取图片流
				XSSFPictureData pic = map.get(key[i]);
				// 获取图片索引
				String picName = key[i].toString();
				// 获取图片格式
				String ext = pic.suggestFileExtension();

				byte[] data = pic.getData();

				FileOutputStream out = new FileOutputStream("D:\\pic" + picName + "." + ext);
				out.write(data);
				out.close();
			}
		}

	}

	public XSSFWorkbook getWorkbook() {
		return workbook;
	}

	public XSSFSheet getSheet() {
		return sheet;
	}

	public int getSheetIndex() {
		return sheetIndex;
	}

	public static void main(String[] args) throws IOException {
		/*InputStream inputStream = new FileInputStream("d:\\success3.xlsx");
		try {
            List<ScmInvoiceImportLog> list = ReadExcel.me(inputStream).setSheetIndex(0).setRowIndex(1).readCoustomTemplet(ScmInvoiceImportLog.class);
            System.out.println(list);
        } catch (InstantiationException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }*/

		// list all demo
		/*List<ArrayList<String>> list = ReadExcel.me(inputStream).setSheetIndex(0).readAllByTitle(4);
		for(ArrayList<String> rowArray : list){
			for(String string : rowArray){
				System.out.print(string+"\t");
			}
			System.out.println();
		}*/
		// read picture demo
		/*List<XSSFPictureData> readPictures = ReadExcel.me(inputStream).readPictures();
		for(XSSFPictureData pictureData:readPictures){
			int type = pictureData.getPictureType();
			if(type == Workbook.PICTURE_TYPE_JPEG){
				File file = new File("d:\\"+System.currentTimeMillis()+".jpg");
				OutputStream outputStream = new FileOutputStream(file);
				outputStream.write(pictureData.getData());
				outputStream.flush();
				outputStream.close();
			}
		}*/
		/*ReadExcel readExcel = ReadExcel.me(inputStream).setSheetIndex(0);
		Map<String, PictureData> dataMap =  readExcel.getSheetPictrues(readExcel.getSheetIndex(),readExcel.getSheet(),readExcel.getWorkbook());
		for(Map.Entry<String,PictureData> entry : dataMap.entrySet()){
			System.out.println("key = "+entry.getKey()+" , value = "+entry.getValue());
		}*/
	}
}
