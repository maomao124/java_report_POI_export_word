package mao;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileOutputStream;

/**
 * Project name(项目名称)：java报表_POI导出word
 * Package(包名): mao
 * Class(类名): Test1
 * Author(作者）: mao
 * Author QQ：1296193245
 * GitHub：https://github.com/maomao124/
 * Date(创建日期)： 2023/6/6
 * Time(创建时间)： 20:53
 * Version(版本): 1.0
 * Description(描述)： 无
 */

public class Test1
{
    public static void main(String[] args)
    {
        XWPFDocument xwpfDocument = new XWPFDocument();
        //创建一个段落
        XWPFParagraph paragraph = xwpfDocument.createParagraph();
        //创建一个片段
        XWPFRun run = paragraph.createRun();
        //设置颜色
        run.setColor("00ff00");
        //大小
        run.setFontSize(18);
        //加粗
        run.setBold(true);
        //文本
        run.setText("Microsoft Office Word是微软公司的一个文字处理器应用程序。");
        //创建一个片段
        run = paragraph.createRun();
        //设置颜色
        run.setColor("0000ff");
        //大小
        run.setFontSize(12);
        //加粗
        run.setBold(false);
        //文本
        run.setText("它最初是由Richard Brodie为了运行DOS的IBM计算机而在1983年编写的。" +
                "随后的版本可运行于Apple Macintosh (1984年)、" +
                "SCO UNIX和Microsoft Windows (1989年)，并成为了Microsoft Office的一部分。" +
                "Word给用户提供了用于创建专业而优雅的文档工具，帮助用户节省时间，并得到优雅美观的结果。");

        //创建一个段落
        paragraph = xwpfDocument.createParagraph();
        //创建一个片段
        run = paragraph.createRun();
        //设置颜色
        run.setColor("ff0000");
        //大小
        run.setFontSize(15);
        //加粗
        run.setBold(true);
        //字体
        run.setFontFamily("宋体");
        //文本
        run.setText("一直以来，Microsoft Office Word 都是最流行的文字处理程序。");
        //创建一个片段
        run = paragraph.createRun();
        //设置颜色
        run.setColor("00ffcc");
        //大小
        run.setFontSize(15);
        //加粗
        run.setBold(true);
        //字体
        run.setFontFamily("黑体");
        //文本
        run.setText("作为 Office 套件的核心程序， Word 提供了许多易于使用的文档创建工具，" +
                "同时也提供了丰富的功能集供创建复杂的文档使用。哪怕只使用 Word 应用一点文本格式化操作或图片处理，" +
                "简单的文档变得比只使用纯文本更具吸引力。\n" +
                "\n");



        try (FileOutputStream fileOutputStream = new FileOutputStream("./out.docx"))
        {
            xwpfDocument.write(fileOutputStream);
            xwpfDocument.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}
