package mao;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.FileOutputStream;

/**
 * Project name(项目名称)：java报表_POI导出word
 * Package(包名): mao
 * Class(类名): Test2
 * Author(作者）: mao
 * Author QQ：1296193245
 * GitHub：https://github.com/maomao124/
 * Date(创建日期)： 2023/6/6
 * Time(创建时间)： 21:14
 * Version(版本): 1.0
 * Description(描述)： 无
 */

public class Test2
{

    /**
     * 得到int随机
     *
     * @param min 最小值
     * @param max 最大值
     * @return int
     */
    public static int getIntRandom(int min, int max)
    {
        if (min > max)
        {
            min = max;
        }
        return min + (int) (Math.random() * (max - min + 1));
    }

    public static void main(String[] args)
    {
        XWPFDocument xwpfDocument = new XWPFDocument();

        createTable(xwpfDocument);

        xwpfDocument.createParagraph().createRun().setText("");

        createTable(xwpfDocument);

        try (FileOutputStream fileOutputStream = new FileOutputStream("./out2.docx"))
        {
            xwpfDocument.write(fileOutputStream);
            xwpfDocument.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    /**
     * 创建表格
     *
     * @param xwpfDocument xwpf文档
     */
    private static void createTable(XWPFDocument xwpfDocument)
    {
        //创建表格
        XWPFTable table = xwpfDocument.createTable(11, 4);
        table.setWidth(5000);
        //创建行
        XWPFTableRow row = table.getRow(0);
        //创建单元格
        XWPFTableCell cell = row.getCell(0);
        //文本
        cell.setText("编号");
        //创建单元格
        cell = row.getCell(1);
        //文本
        cell.setText("姓名");
        //创建单元格
        cell = row.getCell(2);
        //文本
        cell.setText("年龄");
        //创建单元格
        cell = row.getCell(3);
        //文本
        cell.setText("地址");

        for (int i = 1; i < 11; i++)
        {
            //创建行
            row = table.getRow(i);
            //创建单元格
            cell = row.getCell(0);
            //宽度
            //cell.setWidth("12");
            //文本
            cell.setText(String.valueOf(10000 + i));
            //创建单元格
            cell = row.getCell(1);
            //文本
            cell.setText("姓名" + i);
            //创建单元格
            cell = row.getCell(2);
            //文本
            cell.setText(String.valueOf(getIntRandom(15, 30)));
            //创建单元格
            cell = row.getCell(3);
            //文本
            cell.setText("中国");
        }
    }
}
