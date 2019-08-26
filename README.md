# JXLS使用 #
## 1.jxls指令 ##

 1. if指令 
    + `jx:if(condition="employee.payment <= 2000", lastCell="F9", areas=["A9:F9","A18:F18"])`

 2. each指令
    + `jx:each(items="person" var="p" lastCell="C2")`

 3. grid指令(动态表格)
    + `jx:grid(lastCell="A2" headers="headers" data="data" areas=[A1:A1,A2:A2] formatCells="String:A2")`

 4. updateCell指令(自定义处理单元格)

    参考：[jxls-updatecell](http://jxls.sourceforge.net/reference/updatecell_command.html )
    
 5. image指令
    + `jx:image(lastCell="D10" src="image" imageType="PNG")`
    
    + 扩展的
    `jx:image(src="byte[] | JxlsImage | 图片路径（相对图片目录或绝对绝对路径）",lastCell="D10" [,imageType="JPG"] [,size="auto | original"] [,scaleX="1"] [,scaleY="1"])`
 6. **mergeCells** 指令
    + `jx:mergeCells(  lastCell="Merge cell ranges"  [, cols="Number of columns combined"]  [, rows="Number of rows combined"]  [, minCols="Minimum number of columns to merge"]  [, minRows="Minimum number of rows to merge"])`    
      
 7. 公式&& parameeterized formuals
 
    + 普通公式直接在excel单元格中填写(eg. =SUM(E4))
    + 参数化公式 (eg. $[${SUM(E4) * param}])
    + 参考：[官网参数化文档](http://jxls.sourceforge.net/reference/formulas.html) 
    
 8. area listener excel区域块监听
 
   + 参考：[官网实例文档](http://jxls.sourceforge.net/samples/area_listener.html)
   + dafdsafdas

## jxls-reader ##

 1. 编写**.xml
 
 2. java导入excel数据参考Test的importExcel方法