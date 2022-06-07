package demo2;

import com.jacob.activeX.ActiveXComponent;  //下面dispatch类的扩展，为在Microsoft JVM中创建自动化服务提供兼容性
import com.jacob.com.Dispatch; //调度处理类，调度处理类，封装了一些操作来操作office，里面所有的可操作对象基本都是这种类型，所以jacob是一种链式操作模式，就像StringBuilder对象，调用append()方法之后返回的还是StringBuilder对象
import com.jacob.com.Variant;//封装参数数据类型，因为操作office是的一些方法参数，可能是字符串类型，可能是数字类型，虽然都是1，但是不能通过，可以通过Variant来进行转换通用的参数类型.
import com.jacob.com.ComException; //异常类
import com.jacob.com.ComThread;// 线程类初始化com线程，所以会在office转换前、后使用。这个线程好像能力比较弱，可能会有结束不了线程的情况。


////一些方法的说明：

//    •call( )方法：调用COM对象的方法，返回Variant类型值。
//    •invoke( )方法：和call方法作用相同，但是不返回值。

//    •get( )方法：获取COM对象属性，返回variant类型值。
//    •put( )方法：设置COM对象属性

//    Variant对象的toDispatch()方法：将以上方法返回的Variant类型转换为Dispatch，进行下一次链式操作。
//    这些方法有很多重载
//    ComThread.InitSTA(); ComThread.Release();  com线程的初始化和释放操作。

import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.pdf.PdfWriter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;


public class Word2Pdf {
    //三个宏格式值，用来确定目标文件的类型。
    static final int wdDoNotSaveChanges = 0;

    static final int wdFormatPDF = 17;

    static final int ppSaveAsPDF = 32;

    static  int  flag = 0;

    public static void main(String[] args) {

//网络地址
//       String source1 = "https://syr.s3.us-east-1.wasabisys.com/33.xlsx?AWSAccessKeyId=E5S4VAV6YYDYSZ3UKMAD&Expires=1653542057&Signature=vmf%2B%2FQ0iZAFc0CMUHQihfT9Gkgw%3D";
//       String target1 = "C:\\Users\\Administrator\\Desktop\\网络xlsx.pdf";

//本地地址
        Word2Pdf mergeObj = new Word2Pdf();//创建调用四种转换方法的实例对象 mergeObj

        //ppt地址
//        String source1 = "C:\\Users\\Administrator\\Desktop\\jiami.ppt";
//        String target1 = "C:\\Users\\Administrator\\Desktop\\jiami_ppt.pdf";

        String source1 = "C:\\Users\\Administrator\\Desktop\\111.docx";
        String target1 = "C:\\Users\\Administrator\\Desktop\\1111.pdf";
        //excel地址
        String source2 = "C:\\Users\\Administrator\\Desktop\\1.xlsx";
        String target2 = "C:\\Users\\Administrator\\Desktop\\1_xlsx.pdf";
        //doc地址
        String source3 = "C:\\Users\\Administrator\\Desktop\\%s.docx";
        String target3 = "C:\\Users\\Administrator\\Desktop\\%s_docx.pdf";

        //使用String.format 实现将后面的参数拼接到上面的地址。
        String tarUrl= String.format(target1,"2");
        String souUrl= String.format(source1,"2");

        //调用方法，输入源地址和生成地址，第三个参数是加密文件的密码，你不输入也行，不影响。
//        mergeObj.excel2pdf(source1, target1);
        mergeObj.word2pdf(source1, target1);




        //img 转为pdf
//        try {
//            mergeObj.imgToPdf(source1, target1);
//        }catch (Exception e){
//            System.out.println("========Error: Document conversion failed: " + e.getMessage());
//        }
    }


    public void word2pdf(String source, String target) {
        System.out.println("启动Word");
        long start = System.currentTimeMillis();
        ActiveXComponent app = null;
        flag=0;
        try {
            File tofile = new File(target);                                                           //根据target新建一个文件对象
            if (tofile.exists()) {                                                                    //如果已存在，则删除文件。  可以理解如果同路径下有同名PDF，则覆盖掉原来的PDF
                tofile.delete();
            }
            ComThread.InitSTA();
            app = new ActiveXComponent("Word.Application");   //  这里可以和90行合并： ActiveXComponent app = new ActiveXComponent("Word.Application"); 创建一个word的应用对象（word对象要传入“Word.Application” ）
            //  你想创建excel则把参数改成：Excel.Application
            app.setProperty("Visible", false);                          //  设置这个app,也就是word实例的操作的“Visible” 属性的值为false ， 效果就是只在后台操作。 设置的方法是setProperty

            Dispatch docs = app.getProperty("Documents").toDispatch();  // docs 表示 word 的所有文档窗口（ word 是多文档应用程序）；使用getProperty方法，得到app的Documents属性，然后执行toDispatch()将结果变为Dispatch类型；
            System.out.println("Open document" + source);

            String password ="1";


            Dispatch doc = Dispatch.call
                    (docs, "Open", source, false, true,false,new Variant(password))
                    .toDispatch();  //  有了文档对象集合，我们就可以来操作文档了，链式操作就此开始：call方法，调用open方法，传递一个参数，返回一个我们的word文档对象；
            //   五个参数：第一个参数是我们之前获得的文档集，第二个参数是设定操作“Open” 打开，还有“SaveAS”另存为 ，“Close” 关闭。
            //   第三个是文件地址，
            //   第四个是是否进行转换(如果你要转的话，就写上对应的宏格式码。)，
            //   第五个是是否只读。
            //   第六个？
            //   第七个是密码
            //   将返回值送给doc（貌似返回值的内容是定位到哪个文档，执行什么操作）

            System.out.println("Convert document to PDF " + target);

            Dispatch.call(doc, "SaveAs", target, wdFormatPDF, true);                 //call没有返回值，所以直接调用了。 第一个参数是文档，第二个参数是确定操作，第三个是保存的路径，第四个是word 保存为 pdf 格式宏，值为 17(和上面那个是否转换一样)，

            Dispatch.call(doc, "Close");                                        //关闭，去掉这行，也没有影响，包括多个word文档转换的情况下，包括不同文档打印。这个false应该是文件地址。
            long end = System.currentTimeMillis();
            System.out.println("Conversion completed.. Time: " + (end - start) + "ms.");
        } catch (Exception e) {
            System.out.println("========Error: Document conversion failed: " + e.getMessage());
            flag=1;
        } finally {
            if (app != null) {                                                             // 如果为NULL会怎样
                app.invoke("Quit",new Variant[]{});                                  // 加上这个可以关闭office进程
            }
        }
        ComThread.Release();                                                                //这个线程释放，由于我没试过多线程，还不太清除它是否具备多线程稳定运行的能力

    }


    public void ppt2pdf(String source, String target) {
        System.out.println("Start PPT");
        long start = System.currentTimeMillis();
        ActiveXComponent app = null;


        ComThread.InitSTA();
        try {
            app = new ActiveXComponent("Powerpoint.Application");

            Dispatch presentations = app.getProperty("Presentations").toDispatch();

            Dispatch presentation = Dispatch.call(
                    presentations,
                    "Open",
                    source,
                    true,
                    true,
                    false
                    )
                    .toDispatch();  //这里是六个参数，word2PDF是五个参数，不同？

            File tofile = new File(target);
            if (tofile.exists()) {
                tofile.delete();
            }
            Dispatch.call(presentation, "SaveAs", target, ppSaveAsPDF);

            Dispatch.call(presentation, "Close");
            long end = System.currentTimeMillis();


        } catch (Exception e) {
            System.out.println("========Error: Document conversion failed: " + e.getMessage());
        } finally {
            if (app != null) {
                app.invoke("Quit");
            }
        }
        ComThread.Release();
    }

    public void excel2pdf(String source, String target) {
        System.out.println("启动Excel");
        long start = System.currentTimeMillis();
        // start excel(Excel.Application)
        ActiveXComponent app = new ActiveXComponent("Excel.Application");
        ComThread.InitSTA();
        try {

            app.setProperty("Visible", false);
            Dispatch workbooks = app.getProperty("Workbooks").toDispatch();

            System.out.println("Open document" + source);

            Dispatch workbook = Dispatch.invoke(
                    workbooks,
                    "Open",
                    Dispatch.Method,   //dispatch.method是什么？
                    new Object[]{
                            source,
                            new Variant(false),
                            new Variant(true),
                            "1",
                            "pwd"
                    },
                    new int[1]).toDispatch();
            File tofile = new File(target);
            if (tofile.exists()) {
                tofile.delete();
            }
            Dispatch.invoke(
                    workbook,
                    "SaveAs",
                    Dispatch.Method,
                    new Object[]{
                            target,
                            new Variant(57),
                            new Variant(false),
                            new Variant(57),
                            new Variant(57),
                            new Variant(false),
                            new Variant(true),
                            new Variant(57),
                            new Variant(true),
                            new Variant(true),
                            new Variant(true)
                    },
                    new int[1]
            );
            Variant f = new Variant(false);

            System.out.println("Convert document to PDF " + target);

            Dispatch.call(workbook, "Close", f);

            long end = System.currentTimeMillis();

            System.out.println("Conversion completed.. Time: " + (end - start) + "ms.");

        } catch (Exception e) {

            System.out.println("========Error: Document conversion failed: " + e.getMessage());

        } finally {
            if (app != null) {
                app.invoke("Quit", new Variant[]{});
            }
        }
        ComThread.Release();
    }

}
