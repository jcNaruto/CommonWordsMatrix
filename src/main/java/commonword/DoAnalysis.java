package commonword;

/**
 * create by jiacheng on 2019/9/25
 */
public class DoAnalysis {

    public static void main(String[] args) {
        //构造函数中传入知网中导入本地的excel的绝对路径，之后的共词矩阵等也在这里生成
        CommonWordAnalysis commonWordAnalysis = new CommonWordAnalysis("D:\\irshomework\\");
        //第一个参数输入知网中导入本地的excel name
        //第二个参数为输入高频词阈值
        commonWordAnalysis.doAnalysis("jiacheng",10);
    }

}
