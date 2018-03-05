package smartXLS;

import com.smartxls.ChartFormat;
import com.smartxls.ChartShape;
import com.smartxls.WorkBook;

import java.awt.*;

public class CreateChart {

    public static void main(String[] args) {
        WorkBook workBook = new WorkBook();
        try
        {
//set data
            workBook.setText(0, 1, "Jan");
            workBook.setText(0, 2, "Feb");
            workBook.setText(0, 3, "Mar");
            workBook.setText(0, 4, "Apr");
            workBook.setText(0, 5, "Jun");

            workBook.setText(1, 0, "Comfrey");
            workBook.setText(2, 0, "Bananas");
            workBook.setText(3, 0, "Papaya");
            workBook.setText(4, 0, "Mango");
            workBook.setText(5, 0, "Lilikoi");
            for (int col = 1; col <= 5; col++)
                for (int row = 1; row <= 5; row++)
                    workBook.setFormula(row, col, "RAND()");
            workBook.setText(6, 0, "Total");
            workBook.setFormula(6, 1, "SUM(B2:B6)");
            workBook.setSelection("B7:F7");
//auto fill the range with the first cell's formula or data
            workBook.editCopyRight();

            int left = 1;
            int top = 7;
            int right = 13;
            int bottom = 31;

//create chart with it's location
            ChartShape chart = workBook.addChart(left, top, right, bottom);
            chart.setChartType(ChartShape.Column);

            //stacked column chart
            chart.setPlotStacked(true);//set plot serires stacked
            chart.setBarGapRatio(-100);
//link data source, link each series to columns(true to rows).
            chart.setLinkRange("Sheet1!$a$1:$F$6", false);
//set axis title
            chart.setAxisTitle(ChartShape.XAxis, 0, "X-axis data");
            chart.setAxisTitle(ChartShape.YAxis, 0, "Y-axis data");
//set series name
            chart.setSeriesName(0, "My Series number 1");
            chart.setSeriesName(1, "My Series number 2");
            chart.setSeriesName(2, "My Series number 3");
            chart.setSeriesName(3, "My Series number 4");
            chart.setSeriesName(4, "My Series number 5");
            chart.setTitle("My Chart");
//set plot area's color to darkgray
            ChartFormat chartFormat = chart.getPlotFormat();
            chartFormat.setSolid();
            chartFormat.setForeColor(Color.WHITE.getRGB());
            chart.setPlotFormat(chartFormat);

//set series 0's color to blue
            ChartFormat seriesformat = chart.getSeriesFormat(0);
            seriesformat.setSolid();
            seriesformat.setForeColor(Color.BLUE.getRGB());
            chart.setSeriesFormat(0, seriesformat);


//set series 1's color to red
            seriesformat = chart.getSeriesFormat(1);
            seriesformat.setSolid();
            seriesformat.setForeColor(Color.RED.getRGB());
            chart.setSeriesFormat(1, seriesformat);

//set chart title's font property
            ChartFormat titleformat = chart.getTitleFormat();
            titleformat.setFontSize(14*20);
            titleformat.setFontUnderline(true);
            titleformat.setTextRotation(90);
            chart.setTitleFormat(titleformat);

            workBook.writeXLSX("Chart.xlsx");
        }
        catch (Exception ex)
        {
            ex.printStackTrace();
        }
    }
}
