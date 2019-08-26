package com.self.jxlsdemo;

import com.self.jxlsdemo.command.ImageCommand;
import com.self.jxlsdemo.config.JxlsConfig;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.Context;
import org.jxls.transform.poi.PoiContext;
import org.jxls.util.JxlsHelper;

/**
 * jxls excel导出构建器类
 */
public abstract class JxlsBuilder {

  static {
    //注册 jx 命令
    XlsCommentAreaBuilder.addCommandMapping("image", ImageCommand.class);
    //XlsCommentAreaBuilder.addCommandMapping("keep", KeepCommand.class);
    // XlsCommentAreaBuilder.addCommandMapping("grid", GridCommand.class);
  }

  private JxlsHelper jxlsHelper = JxlsHelper.getInstance();

  private Context context;
  private InputStream in;//模板输入流
  private OutputStream out;//excel导出流
  private TemplateConfig config;//导出配置

  private JxlsBuilder() {
    //可以初始化为poiContext
    context = new PoiContext();
    //context = new Context();
    //默认配置
    config = new TemplateConfig();
  }

  private JxlsBuilder(InputStream in) {
    this();
    this.in = in;
  }

  private JxlsBuilder(File inFile) {
    this();
    if (!inFile.exists()) {
      throw new IllegalArgumentException("模板文件不存在：" + inFile.getAbsolutePath());
    }
    if (!inFile.getName().toLowerCase().endsWith("xls") &&
        !inFile.getName().toLowerCase().endsWith("xlsx")) {
      throw new IllegalArgumentException("不支持非excel文件：" + inFile.getName());
    }
    try {
      in = new FileInputStream(inFile);
    } catch (FileNotFoundException e) {
      throw new IllegalArgumentException("文件读取失败：" + inFile.getAbsolutePath(), e);
    }
  }

  /**
   * @param in 模板文件流
   */
  public static JxlsBuilder getBuilder(InputStream in) {
    return new JxlsBuilderImpl(in);
  }

  /**
   * @param templateFile 模板文件地址
   */
  public static JxlsBuilder getBuilder(File templateFile) {
    return new JxlsBuilderImpl(templateFile);
  }

  /**
   * @param filePath 模板文件路径，可以是绝对路径，也可以是模板存放目录的文件名
   */
  public static JxlsBuilder getBuilder(String filePath) {
    //判断是相对路径还是绝对路径
    if (!JxlsUtil.me().isAbsolutePath(filePath)) {
      if (JxlsConfig.getTemplateRoot().startsWith("classpath:")) {
        //文件在jar包内
        String templateRoot = JxlsConfig.getTemplateRoot().replaceFirst("classpath:", "");
        InputStream resourceAsStream = JxlsBuilder.class.getResourceAsStream(templateRoot + "/" + filePath);
        return new JxlsBuilderImpl(resourceAsStream);
      } else {
        //相对路径就从模板目录获取文件
        return new JxlsBuilderImpl(new File(JxlsConfig.getTemplateRoot() +
            File.separator + filePath));
      }
    } else {
      //绝对路径
      return new JxlsBuilderImpl(new File(filePath));
    }
  }

  public JxlsHelper getJxlsHelper() {
    return jxlsHelper;
  }

  /**
   * 生成excel文件
   */
  public JxlsBuilder build() throws Exception {
    if (JxlsUtil.me().hasText(config.getImageRoot())) {
      context.putVar("_imageRoot", config.getImageRoot());
    }
    context.putVar("_ignoreImageMiss", config.ignoreImageMiss);
    if (config.isGridTemplate) {
      if (JxlsUtil.me().hasText(config.targetCell)) {
        jxlsHelper.processGridTemplateAtCell(in, out, context, config.objectProps, config.targetCell);
        return this;
      }
      jxlsHelper.processGridTemplate(in, out, context, config.objectProps);
    } else {
      if (JxlsUtil.me().hasText(config.targetCell)) {
        jxlsHelper.processTemplateAtCell(in, out, context, config.targetCell);
        return this;
      }
      jxlsHelper.processTemplate(in, out, context);
    }
    return this;
  }


  private static class JxlsBuilderImpl extends JxlsBuilder {

    public JxlsBuilderImpl(InputStream in) {
      super(in);
    }

    public JxlsBuilderImpl(File inFile) {
      super(inFile);
    }
  }

  /**
   * 配置动态表格模板
   *
   * @param objectProps (data数据为Collection<Object> 类型必须设置)
   */
  public JxlsBuilder configDynamicGrid(String objectProps) {
    this.config.objectProps = objectProps;
    this.config.isGridTemplate = true;
    return this;
  }

  /**
   * 配置导出目标sheet页（不配置，默认导出到模板所在的sheet页）
   */
  public JxlsBuilder configTargetCell(String targetCell) {
    this.config.targetCell = targetCell;
    return this;
  }

  /**
   * 自定义配置
   */
  public JxlsBuilder customConfig(TemplateConfig config) {
    this.config = config;
    return this;
  }

  private static class TemplateConfig {

    private Map<String, Object> funcs;
    private String imageRoot = JxlsConfig.getImageRoot();
    private boolean ignoreImageMiss = false;
    private boolean isGridTemplate = false;//是否为动态的表格模板
    private String objectProps;//动态表格的对象属性字符串(comma separated)
    private String targetCell;//excel导出的目标sheet页

    public Map<String, Object> getFuncs() {
      return funcs;
    }

    public TemplateConfig setFuncs(Map<String, Object> funcs) {
      this.funcs = funcs;
      return this;
    }

    public String getImageRoot() {
      return imageRoot;
    }

    public TemplateConfig setImageRoot(String imageRoot) {
      this.imageRoot = imageRoot;
      return this;
    }

    public boolean isIgnoreImageMiss() {
      return ignoreImageMiss;
    }

    public TemplateConfig setIgnoreImageMiss(boolean ignoreImageMiss) {
      this.ignoreImageMiss = ignoreImageMiss;
      return this;
    }


    public boolean isGridTemplate() {
      return isGridTemplate;
    }

    public TemplateConfig setGridTemplate(boolean gridTemplate) {
      isGridTemplate = gridTemplate;
      return this;
    }

    public String getObjectProps() {
      return objectProps;
    }

    public TemplateConfig setObjectProps(String objectProps) {
      this.objectProps = objectProps;
      return this;
    }

    public String getTargetCell() {
      return targetCell;
    }

    public TemplateConfig setTargetCell(String targetCell) {
      this.targetCell = targetCell;
      return this;
    }

  }


  /**
   * 指定输出流
   */
  public JxlsBuilder out(OutputStream out) {
    this.out = out;
    return this;
  }

  /**
   * 指定输出文件
   */
  public JxlsBuilder out(File outFile) throws Exception {
    this.out = new FileOutputStream(outFile);
    return this;
  }

  /**
   * 指定输出文件绝对路径
   */
  public JxlsBuilder out(String outPath) throws Exception {
    File file = new File(outPath);
    this.out = new FileOutputStream(file);
    return this;
  }

  /**
   * 添加数据
   */
  public JxlsBuilder putVar(String name, Object value) {
    context.putVar(name, value);
    return this;
  }

  /**
   * 添加数据
   */
  public JxlsBuilder putAll(Map<String, Object> map) {
    for (String key : map.keySet()) {
      putVar(key, map.get(key));
    }
    return this;
  }


  /**
   * 删除数据
   */
  public JxlsBuilder removeVar(String name) {
    context.removeVar(name);
    return this;
  }

  /**
   * 获取数据
   */
  public Object getVar(String name) {
    return context.getVar(name);
  }


  public TemplateConfig getConfig() {
    return config;
  }

  public JxlsBuilder setConfig(TemplateConfig config) {
    this.config = config;
    return this;
  }


}
