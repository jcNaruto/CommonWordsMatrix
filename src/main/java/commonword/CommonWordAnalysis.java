package commonword;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * @Author jiacheng
 * @Description  关键词共词矩阵分析
 * @Date 18:15 2019/9/26
 * @Param
 * @throws
 * @return
 */
public class CommonWordAnalysis {

	private  Map<String,String> bucket = new LinkedHashMap<String, String>();//key为文章id,value为该文章关键词，用于形成共词矩阵
	private  Map<String,Integer> countBucket = new LinkedHashMap<String,Integer>();//key为关键词，value为该词出现的次数
	private  Map<String,Integer> countBucket2 = new LinkedHashMap<String, Integer>();//经过用户指定高频词筛选之后的countBucket
	private  List<Integer> totalValue = new ArrayList<Integer>();//顺序存储了共词矩阵每个单元格的content
	private String filePath;

	public CommonWordAnalysis(String filePath){
		this.filePath = filePath;
	}


	public  void doAnalysis(String fileName, int highFrequencyWordthreshold){
		try {
			//输入知网中导入本地的excel name
			readExcel(fileName);
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		try {
			getWord();
		} catch (Exception e) {
			e.printStackTrace();
		}
		//输入高频词阈值
		getHighF(highFrequencyWordthreshold);

		try {
            getTotalWordMatrix();

			writeExcelCommandWordMatrix();
			writeExcelSimilarMatrix();
			writeExcelDifferentMatirx();

			getWordTop();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (WriteException e) {
			e.printStackTrace();
		}catch (Exception e) {
			e.printStackTrace();
		}
	}

    public void readExcelToExclude(String cName) throws BiffException, IOException {
        Set<String> word = new HashSet<String>();
        InputStream stream = new FileInputStream(filePath + "jiacheng" + ".xls");
        // 获得工作簿对象
        Workbook workbook = Workbook.getWorkbook(stream);
        //获得第一张工作表
        Sheet sheet = workbook.getSheet(0);
        // 遍历工作表
        // 获总行数
        int rows = sheet.getRows();
        // 获得总列数
        int cols = sheet.getColumns();

        int tempConfirmCol = -1;
        for (int col = 0; col < cols; col++) {
            if (cName.equals(sheet.getCell(col, 0).getContents())) {
                tempConfirmCol = col;
            }
        }
        int count = 0;
        for (int row = 1; row < rows; row++) {
			if (tempConfirmCol != -1) {
				String keyWords = sheet.getCell(tempConfirmCol, row).getContents();
				if ("".equals(keyWords) || keyWords == null) {
					System.out.println( cName+"缺失的论文为" + row + "篇");
					count++;
				}
			}
		}
			System.out.println(cName+"列不合格论文为" + count + "篇");
            workbook.close();
        }





    public void readExcel(String name) throws BiffException, IOException {
    	Set<String> word = new HashSet<String>();
    	InputStream stream = new FileInputStream(filePath +name+".xls");
        // 获得工作簿对象
        Workbook workbook = Workbook.getWorkbook(stream);
        //获得第一张工作表
        Sheet sheet = workbook.getSheet(0);
        // 遍历工作表
     	// 获总行数
        int rows = sheet.getRows();
        // 获得总列数
        int cols =sheet.getColumns();

        int tempConfirmCol = -1;
        for(int col = 0; col < cols; col++){
			if("Keyword-关键词".equals(sheet.getCell(col, 0).getContents())){
				tempConfirmCol = col;
			}
		}
		for(int row = 1; row < rows; row++){
			if(tempConfirmCol != -1){
				//keyWords单元格content
				String keyWords = sheet.getCell(tempConfirmCol, row).getContents();
				String[] split = keyWords.split(";;");
				for(String keyWord : split){
					if("".equals(keyWord)){
						continue;
					}
					if(countBucket.containsKey(keyWord)){
						Integer count  = countBucket.get(keyWord);
						countBucket.put(keyWord,count+1);
					}else{
						countBucket.put(keyWord, 1);
					}
				}
				bucket.put(String.valueOf(row),";;"+keyWords+";;");
				System.out.println(row + "篇论文读取完毕" );
			}
		}
         workbook.close();
     }

     /**
      * @Author Jiacheng
      * @Description 得到所有词
      * @Date 8:09 2019/7/11
      * @Param []
      * @throws
      * @return void
      */
     public void getWord() throws Exception{
		 Set<Map.Entry<String, Integer>> entries = countBucket.entrySet();

		 File xlsFile = new File("D:\\irshomework\\AllWords.xls");
		 // 创建一个工作簿
		 WritableWorkbook workbook = Workbook.createWorkbook(xlsFile);
		 // 创建一个工作表
		 WritableSheet sheet = workbook.createSheet("sheet1", 0);

		 int c = 0;
		 int r = 0;
		for(Map.Entry<String, Integer> entry : entries){
			entry.getValue();
			if(r == 200){
				r = 0;
				c++;
			}
			sheet.addCell(new Label(c, r++, entry.getKey()));
			sheet.addCell(new Label(c, r++, String.valueOf(entry.getValue())));
		}
		 workbook.write();
		 workbook.close();
	 }


	public void getWordTop() throws Exception{
		Set<Map.Entry<String, Integer>> entries = countBucket2.entrySet();

		File xlsFile = new File("D:\\irshomework\\WordTop.xls");
		// 创建一个工作簿
		WritableWorkbook workbook = Workbook.createWorkbook(xlsFile);
		// 创建一个工作表
		WritableSheet sheet = workbook.createSheet("sheet1", 0);

		int c = 0;
		int r = 0;
		for(Map.Entry<String, Integer> entry : entries){
			entry.getValue();
			if(r == 10){
				r = 0;
				c++;
			}
			sheet.addCell(new Label(c, r++, entry.getKey()));
			sheet.addCell(new Label(c, r++, String.valueOf(entry.getValue())));
		}
		workbook.write();
		workbook.close();
	}


     /**
      * @Author Jiacheng
      * @Description 得到高频词
      * @Date 16:21 2019/7/8
      * @Param []
      * @throws
      * @return void
      */
     public void getHighF(int f){
		for(Map.Entry<String,Integer> entry : countBucket.entrySet()){
			if(entry.getValue() > f && !entry.getKey().equals("")){
				countBucket2.put(entry.getKey(),entry.getValue());
			}
		}
	 }


    /**
     * @Author Jiacheng
     * @Description 得到共词矩阵
     * @Date 9:52 2019/7/9
     * @Param []
     * @throws
     * @return void
     */
	 public void getTotalWordMatrix(){
         List<String> array = new ArrayList<String>(countBucket2.keySet());//共词矩阵的行列有顺序
         for(int i = 0; i < array.size(); i++){
             for(int j = 0; j < array.size(); j++){
                 int tempValue = 0;
                 String temp1 =";;"+ array.get(i) + ";;";
                 String temp2 =";;"+  array.get(j) + ";;";
                 for(int k = 0;k < bucket.entrySet().size(); k++){//遍历文章数组
                     String str = bucket.get(String.valueOf(k + 1));
                     if(str.contains(temp1) && str.contains(temp2)){
                     	if(temp2.equals(";;网络舆情;;") && temp1.equals(";;网络舆情;;")){
							System.out.println("==============" + str +"========="+ tempValue);
						}
                         tempValue++;
                     }
                 }
                 //两个词共同出现在几篇文章中，顺序是依次递增的
                 totalValue.add(tempValue);
             }
         }
     }





     /**
      * @Author Jiacheng
      * @Description 写关键词共词矩阵
      * @Date 8:15 2019/7/10
      * @Param []
      * @throws
      * @return void
      */
    public void writeExcelCommandWordMatrix()throws IOException, WriteException{
		Set<String> strings = countBucket2.keySet();
		List<String> temps = new ArrayList<String>(strings);
		File xlsFile = new File("D:\\irshomework\\CommandWordMatrix.xls");
		// 创建一个工作簿
		WritableWorkbook workbook = Workbook.createWorkbook(xlsFile);
		// 创建一个工作表
		WritableSheet sheet = workbook.createSheet("sheet1", 0);
		for (int row = 0; row < countBucket2.entrySet().size(); row++) {
			sheet.addCell(new Label(0, row + 1, temps.get(row)));
			sheet.addCell(new Label(row + 1, 0, temps.get(row)));
		}
		int currIndex = 0;
		for(int col = 1; col <= countBucket2.size(); col++){
			for(int r = 1 ; r <= countBucket2.size(); r++){
				sheet.addCell(new Label(col, r, String.valueOf(totalValue.get(currIndex++))));
			}
		}
		workbook.write();
		workbook.close();
	}

	/**
	 * @Author Jiacheng
	 * @Description 写关键词相似矩阵，代表了两个关键词之间的亲疏关系
	 * @Date 20:52 2019/9/25
	 * @Param []
	 * @throws
	 * @return void
	 */
	public void writeExcelSimilarMatrix()throws IOException, WriteException{
		Set<String> strings = countBucket2.keySet();
		List<String> temps = new ArrayList<String>(strings);
		File xlsFile = new File("D:\\irshomework\\SimilarMatrix.xls");
		// 创建一个工作簿
		WritableWorkbook workbook = Workbook.createWorkbook(xlsFile);
		// 创建一个工作表
		WritableSheet sheet = workbook.createSheet("sheet1", 0);
		for (int row = 0; row < countBucket2.entrySet().size(); row++) {
			sheet.addCell(new Label(0, row + 1, temps.get(row)));
			sheet.addCell(new Label(row + 1, 0, temps.get(row)));
		}
		Set<Map.Entry<String, Integer>> entries = countBucket2.entrySet();
		List<Map.Entry<String, Integer>> list = new ArrayList<Map.Entry<String, Integer>>(entries);
		int currIndex = 0;
		for(int col = 1; col <= countBucket2.size(); col++){
			for(int r = 1 ; r <= countBucket2.size(); r++){
				Integer integer = totalValue.get(currIndex++);//AB两个词同时出现的次数
				Integer value = list.get(col - 1).getValue();//col上对应的词出现的次数
				Integer value1 = list.get(r - 1).getValue();//row上对应的词出现的次数

				if(integer.equals(value) && value.equals(value1)){
					sheet.addCell(new Label(col, r, String.valueOf(1)));
				}else{
					double v = 1.0 * integer / (Math.sqrt(value) * Math.sqrt(value1));
					sheet.addCell(new Label(col, r, String.valueOf(v)));
				}
			}
		}
		workbook.write();
		workbook.close();
	}


	/**
	 * @Author Jiacheng
	 * @Description  写关键词相异矩阵，数值越小说明两个关键词距离越近，在聚类中越容易被当作一类
	 * @Date 20:53 2019/9/25
	 * @Param []
	 * @throws
	 * @return void
	 */
	public void writeExcelDifferentMatirx()throws IOException, WriteException{
		Set<String> strings = countBucket2.keySet();
		List<String> temps = new ArrayList<String>(strings);
		File xlsFile = new File("D:\\irshomework\\DifferentMatirx.xls");
		// 创建一个工作簿
		WritableWorkbook workbook = Workbook.createWorkbook(xlsFile);
		// 创建一个工作表
		WritableSheet sheet = workbook.createSheet("sheet1", 0);
		for (int row = 0; row < countBucket2.entrySet().size(); row++) {
			sheet.addCell(new Label(0, row + 1, temps.get(row)));
			sheet.addCell(new Label(row + 1, 0, temps.get(row)));
		}
		Set<Map.Entry<String, Integer>> entries = countBucket2.entrySet();
		List<Map.Entry<String, Integer>> list = new ArrayList<Map.Entry<String, Integer>>(entries);
		int currIndex = 0;
		for(int col = 1; col <= countBucket2.size(); col++){
			for(int r = 1 ; r <= countBucket2.size(); r++){
				Integer integer = totalValue.get(currIndex++);//AB两个词同时出现的次数
				Integer value = list.get(col - 1).getValue();//col上对应的词出现的次数
				Integer value1 = list.get(r - 1).getValue();//row上对应的词出现的次数

				if(integer.equals(value) && value.equals(value1)){
					sheet.addCell(new Label(col, r, String.valueOf(0)));
				}else{
					double v = 1.0 * integer / (Math.sqrt(value) * Math.sqrt(value1));
					sheet.addCell(new Label(col, r, String.valueOf(1-v)));
				}
			}
		}
		workbook.write();
		workbook.close();
	}

    
}
					