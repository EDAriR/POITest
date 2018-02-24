package chart;

public class OOXMLBarChart extends OOXMLChart {

    /**
     * @param args
     */
    public static void main(String[] args) {
        String inFile = "tpl_bar.xlsx"; // 模板文件
        String outFile = "111out_bar.xlsx"; // 输出文件
        // 标题项目
        String[] titles = new String[] { "xxx项目一", "项sss目二", "项目xx"};
        // 数据
        double[] values = new double[] { 150, 20, 30, 87, 94 };
        System.out.println("读取模板 " + inFile);
        createChart(titles, values, inFile, outFile);
        System.out.println("输出文件 " + outFile);
    }
}
