package demo2;

import com.jacob.activeX.ActiveXComponent;  //下面dispatch类的扩展，为在Microsoft JVM中创建自动化服务提供兼容性
import com.jacob.com.Dispatch; //调度处理类，调度处理类，封装了一些操作来操作office，里面所有的可操作对象基本都是这种类型，所以jacob是一种链式操作模式，就像StringBuilder对象，调用append()方法之后返回的还是StringBuilder对象
import com.jacob.com.Variant;//封装参数数据类型，因为操作office是的一些方法参数，可能是字符串类型，可能是数字类型，虽然都是1，但是不能通过，可以通过Variant来进行转换通用的参数类型
import com.jacob.com.ComException; //异常类
import com.jacob.com.ComThread;// 线程类初始化com线程，所以会在office转换前、后使用。


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
//三个宏格式值，用来确定转换源文件和结果文件的类型。
    static final int wdDoNotSaveChanges = 0;

    static final int wdFormatPDF = 17;

    static final int ppSaveAsPDF = 32;

    public static void main(String[] args) {

//网络地址
//      String source1 = "http://localhost:3000/public/120mb.docx";
//       String target1 = "C:\\Users\\Administrator\\Desktop\\800kb.pdf";


//本地地址
        String source1 = "C:\\Users\\Administrator\\Desktop\\800kb.docx";
        String target1 = "C:\\Users\\Administrator\\Desktop\\800kb.pdf";


        String source2 = "C:\\Users\\Administrator\\Desktop\\ppt_test2.ppt";
        String target2 = "C:\\Users\\Administrator\\Desktop\\ppt_test2.pdf";

        String source3 = "C:\\Users\\Administrator\\Desktop\\15mb.docx";
        String target3 = "C:\\Users\\Administrator\\Desktop\\15mb.pdf";

//word to pdf
        Word2Pdf.word2pdf(source1, target1);
//        Word2Pdf.word2pdf(source3, target3);
//ppt to pdf
        Word2Pdf mergeObj = new Word2Pdf();


        mergeObj.ppt2pdf(source2, target2);

        // excel 转为PPT
//        mergeObj.excel2pdf(source1, target1);

        //img 转为pdf
//        try {
//            mergeObj.imgToPdf(source1, target1);
//        }catch (Exception e){
//            System.out.println("========Error: Document conversion failed: " + e.getMessage());
//        }

//
    }

    public static void word2pdf(String source, String target) {
        System.out.println("启动Word");
        long start = System.currentTimeMillis();
        ActiveXComponent app = null;
        try {

            app = new ActiveXComponent("Word.Application");   //  这里可以和90行合并： ActiveXComponent app = new ActiveXComponent("Word.Application"); 创建一个word的应用对象（word对象要传入“Word.Application” ）
                                                                        //  你想创建excel则把参数改成：Excel.Application
            app.setProperty("Visible", false);                          //  设置这个app,也就是word实例的操作的“Visible” 属性的值为false ， 效果就是只在后台操作。 设置的方法是setProperty

            Dispatch docs = app.getProperty("Documents").toDispatch();  // docs 表示 word 的所有文档窗口（ word 是多文档应用程序）；使用getProperty方法，得到app的Documents属性，然后执行toDispatch()将结果变为Dispatch类型；
            System.out.println("Open document" + source);
            Dispatch doc = Dispatch.call( docs, "Open",source, false, true).toDispatch();  //  有了文档对象集合，我们就可以来操作文档了，链式操作就此开始：call方法，调用open方法，传递一个参数，返回一个我们的word文档对象；
                                                                                                           //   五个参数：第一个参数是我们之前获得的文档集，第二个参数是设定操作“Open” 打开，还有“SaveAS”另存为 ，“Close” 关闭。 第三个是文件地址，第四个是是否进行转换(如果你要转的话，就写上对应的宏格式码。)，第五个是是否只读。
                                                                                                           //   将返回值送给doc（貌似返回值的内容是定位到哪个文档，执行什么操作）

            System.out.println("Convert document to PDF " + target);
            File tofile = new File(target);                                                           //根据target新建一个文件对象
            if (tofile.exists()) {                                                                    //如果已存在，则删除文件。  可以理解如果同路径下有同名PDF，则覆盖掉原来的PDF
                tofile.delete();
            }
            Dispatch.call(doc, "SaveAs", target, wdFormatPDF ,true);                 //call没有返回值，所以直接调用了。 第一个参数是文档，第二个参数是确定操作，第三个是保存的路径，第四个是word 保存为 pdf 格式宏，值为 17(和上面那个是否转换一样)，

            Dispatch.call(doc, "Close", false);                                       //关闭，去掉这行，也没有影响，包括多个word文档转换的情况下，包括不同文档打印。这个false应该是文件地址。
            long end = System.currentTimeMillis();
            System.out.println("Conversion completed.. Time: " + (end - start) + "ms.");
        } catch (Exception e) {
            System.out.println("========Error: Document conversion failed: " + e.getMessage());
        } finally {
            if (app != null) {                                                             // 如果为NULL会怎样
                app.invoke("Quit", wdDoNotSaveChanges);                      // invoke调用com组件？
            }
        }
    }

    public void ppt2pdf(String source, String target) {
        System.out.println("Start PPT");
        long start = System.currentTimeMillis();
        ActiveXComponent app = null;
        try {
            app = new ActiveXComponent("Powerpoint.Application");

            Dispatch presentations = app.getProperty("Presentations").toDispatch();//ok

            Dispatch presentation = Dispatch.call(presentations, "Open", source, true, true, false).toDispatch();  //这里是六个参数，word2PDF是五个参数，不同？


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
    }

    public void excel2pdf(String source, String target) {
        System.out.println("启动Excel");
        long start = System.currentTimeMillis();
        // start excel(Excel.Application)
        ActiveXComponent app = new ActiveXComponent("Excel.Application");
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
                            new Variant(false)
                    },
                    new int[3]).toDispatch();

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
    }

    public boolean imgToPdf(String imgFilePath, String pdfFilePath) throws IOException {
        File file = new File(imgFilePath);
        if (file.exists()) {
            Document document = new Document();
            FileOutputStream fos = null;
            try {
                fos = new FileOutputStream(pdfFilePath);
                PdfWriter.getInstance(document, fos);

                // Add some information about the PDF document, such as author, subject, etc.
                document.addAuthor("arui");
                document.addSubject("test pdf.");
                // set the size of the document
                document.setPageSize(PageSize.A4);
                // open the document
                document.open();
                // write a text
                // document.add(new Paragraph("JUST TEST ..."));
                // read an image
                Image image = Image.getInstance(imgFilePath);
                float imageHeight = image.getScaledHeight();
                float imageWidth = image.getScaledWidth();
                int i = 0;
                while (imageHeight > 500 || imageWidth > 500) {
                    image.scalePercent(100 - i);
                    i++;
                    imageHeight = image.getScaledHeight();
                    imageWidth = image.getScaledWidth();
                    System.out.println("imageHeight->" + imageHeight);
                    System.out.println("imageWidth->" + imageWidth);
                }

                image.setAlignment(Image.ALIGN_CENTER);
                // //Set the absolute position of the image
                // image.setAbsolutePosition(0, 0);
                // image.scaleAbsolute(500, 400);
                // insert an image
                document.add(image);
            } catch (DocumentException de) {
                System.out.println(de.getMessage());
            } catch (IOException ioe) {
                System.out.println(ioe.getMessage());
            }
            document.close();
            fos.flush();
            fos.close();
            return true;
        } else {
            return false;
        }
    }
}