import com.ecnice.oa.vo.wf.api.CommentVo;
import com.mhome.tools.pager.PagerModel;
import com.mhome.tools.pager.Query;
import com.ys.portal.model.news.*;
import com.ys.portal.service.news.*;
import com.ys.tools.common.DateUtil;
import com.ys.tools.common.ReadProperty;
import freemarker.template.Configuration;
import freemarker.template.Template;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Pattern;

/**
 * @Title:导出公告工具类
 * @Description:
 * @Author:YZ
 * @Since:2018年6月25日13:24:44
 * @Version:1.1.0
 * @Copyright:Copyright (c) 浙江蘑菇加电子商务有限公司 2015 ~ 2016 版权所有
 */
public class ExportNoticeUtil {


    private static Pattern FilePattern = Pattern.compile("[\\\\/:*?\"<>|]");

    /**
     * 基于freemarker导出包含富文本的word
     * @param ids
     * @param sessionId
     * @param typeId
     * @param nameSpace
     * @param newsNoticeService
     * @param newsFileService
     * @param response
     * @param readProperty
     * @param basePath
     * @param newsNoticeProcessService
     * @param newsPublishRangeService
     */

    public static void ExportNoticeWord(String ids, String sessionId, String typeId, String nameSpace, INewsNoticeService newsNoticeService,
                                        INewsFileService newsFileService, HttpServletResponse response, ReadProperty readProperty, String basePath, INewsNoticeProcessService newsNoticeProcessService, INewsPublishRangeService newsPublishRangeService){
        try {

            Configuration configuration = new Configuration();
            configuration.setDefaultEncoding("UTF-8");
            String idArr[] = ids.split(",");
            String singlrNoticeName = "";
            for(int i=0;i<idArr.length;i++){
                //下载word
                NewsNotice notice = newsNoticeService.getNoticeById(idArr[i]);
                Map<String, Object> dataMap = ExportNoticeUtil.getData(notice,newsNoticeProcessService,newsPublishRangeService);
                WordHtmlGeneratorHelper.handleAllObject(dataMap);
                String filePath = ExportNoticeUtil.class.getClassLoader().getResource("noticeExportWord.ftl").getPath();
                configuration.setDirectoryForTemplateLoading(new File(filePath).getParentFile());//模板文件所在路径
                Template t = null;
                t = configuration.getTemplate("noticeExportWord.ftl","UTF-8"); //获取模板文件
                File outFile = null;

                if(idArr.length == 1){
                    singlrNoticeName = notice.getArticleNo();
                }
                String formatTitle = StringUtils.isNotBlank(notice.getTitle())?FilePattern.matcher(notice.getTitle()).replaceAll("")
                        :"标题为空";
                outFile = idArr.length == 1?new File(basePath+notice.getArticleNo(),formatTitle+".doc")
                        :new File(basePath+"公文批导"+getTime("yyyyMMdd")+"/"+notice.getArticleNo(),formatTitle+".doc");
                if(!outFile.exists()){
                    outFile.getParentFile().mkdirs();
                }
                Writer out = null;
                out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile),"UTF-8"));
                t.process(dataMap, out); //将填充数据填入模板文件并输出到目标文件
                out.close();

                //下载附件
                NewsFile newsFile = new NewsFile();
                newsFile.setRefId(idArr[i]);
                PagerModel<NewsFile> pm = null;
                if (StringUtils.isNotBlank(newsFile.getRefId())) {
                    pm = newsFileService.getPagerModelByQuery(newsFile, new Query());
                }
                List<NewsFile> files = pm.getRows();

                for(NewsFile file:files){
                    File downLoadFile = idArr.length == 1?new File(basePath+notice.getArticleNo()+"/"+notice.getArticleNo()+"附件",file.getFileName())
                            :new File(basePath+"公文批导"+getTime("yyyyMMdd")+"/"+notice.getArticleNo()+"/"+notice.getArticleNo()+"附件",file.getFileName());
                    downLoadFile(file,readProperty,downLoadFile);
                }
            }

            //压缩文件
            String zipFileName = idArr.length == 1?basePath + singlrNoticeName+".zip":basePath+"公文批导"+getTime("yyyyMMdd")+".zip";
            String FileName = idArr.length == 1?basePath + singlrNoticeName:basePath+"公文批导"+getTime("yyyyMMdd");
            ZipCompressor zc = new ZipCompressor(zipFileName);
            File[] srcfile = new File[1];
            srcfile[0] = new File(FileName);
            zc.compress(srcfile);

            //导出压缩包
            String title = idArr.length == 1?singlrNoticeName+".zip":"公文批导"+getTime("yyyyMMdd")+".zip";
            OutputStream output = response.getOutputStream();
            response.reset();
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + new String(title.getBytes("UTF-8"), "ISO8859-1"));
            response.setContentType("application/octet-stream;charset=UTF-8");

            FileInputStream inStream = new FileInputStream(idArr.length == 1?basePath+singlrNoticeName+".zip":basePath+"公文批导"+getTime("yyyyMMdd")+".zip");
            byte[] buf = new byte[4096];
            int readLength;
            while (((readLength = inStream.read(buf)) != -1)) {
                output.write(buf, 0, readLength);
            }

            inStream.close();
            output.flush();
            output.close();


            String delDir= idArr.length == 1?basePath+singlrNoticeName
                        :basePath+"公文批导"+getTime("yyyyMMdd");
            String delFile= idArr.length == 1?basePath+singlrNoticeName+".zip"
                    :basePath+"公文批导"+getTime("yyyyMMdd")+".zip";

//            //删除临时文件
            DeleteFileUtil.deleteDirectory(delDir);
            DeleteFileUtil.delete(delFile);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 下载附件
     * @param file
     * @param readProperty
     * @param downLoadFile
     */
    private static void downLoadFile(NewsFile file, ReadProperty readProperty,File downLoadFile){
        String path = file.getFilePath();
        BufferedInputStream bis = null;
        InputStream is = null;
        OutputStream os = null;
        try {
            String ftpHost = readProperty.getValue("ftp.host");
            URL url = new URL(path.startsWith("http:")||path.startsWith("https:")?path:ftpHost+path);
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();//利用HttpURLConnection对象,我们可以从网络中获取网页数据.
            conn.setDoInput(true);
            conn.connect();
            is = conn.getInputStream(); //得到网络返回的输入流
            byte[] buffer = new byte[1024];
            bis = new BufferedInputStream(is);

            if(!downLoadFile.exists()){
                downLoadFile.getParentFile().mkdirs();
            }
            os = new DataOutputStream(new FileOutputStream(downLoadFile));
            int j = bis.read(buffer);
            while (j != -1) {
                os.write(buffer, 0, j);
                j = bis.read(buffer);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (bis != null) {
                try {
                    is.close();
                    bis.close();
                    os.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * 获取当前特定格式时间
     * @return
     */
    private static String getTime(String pattern){
        return new SimpleDateFormat(pattern).format(new Date());
    }

    /**
     * 封装数据
     * @param notice
     * @param newsNoticeProcessService
     * @param newsPublishRangeService
     * @return
     */
    private static Map<String, Object> getData(NewsNotice notice, INewsNoticeProcessService newsNoticeProcessService, INewsPublishRangeService newsPublishRangeService) {
        Map<String, Object> res = new HashMap<String, Object>();
        try{
            Class<NewsNotice> clazz = NewsNotice.class;
            Field[] fields = clazz.getDeclaredFields();//遍历字段
            for(Field field:fields){
                field.setAccessible(true);
                String key = field.getName();
                Object val = field.get(notice);
                val = null==val?"":val;

                if("content".equals(key)){//处理富文本
                    RichHtmlHandler handler = new RichHtmlHandler(val.toString());
                    handler.setDocSrcLocationPrex("file:///C:/D1745AB2");
                    handler.setDocSrcParent("test2.files");
                    handler.setNextPartId("01D40A22.6DCACC80");
                    handler.setShapeidPrex("_x56fe__x7247__x0020");
                    handler.setSpidPrex("_x0000_i");
                    handler.setTypeid("#_x0000_t75");

                    handler.handledHtml(false);

                    String bodyBlock = handler.getHandledDocBodyBlock();
                    System.out.println("bodyBlock:\n"+bodyBlock);

                    String handledBase64Block = "";
                    if (handler.getDocBase64BlockResults() != null
                            && handler.getDocBase64BlockResults().size() > 0) {
                        for (String item : handler.getDocBase64BlockResults()) {
                            handledBase64Block += item + "\n";
                        }
                    }
                    if(StringUtils.isBlank(handledBase64Block)){
                        handledBase64Block = "";
                    }
                    res.put("imagesBase64String", handledBase64Block);

                    String xmlimaHref = "";
                    if (handler.getXmlImgRefs() != null
                            && handler.getXmlImgRefs().size() > 0) {
                        for (String item : handler.getXmlImgRefs()) {
                            xmlimaHref += item + "\n";
                        }
                    }

                    if(StringUtils.isBlank(xmlimaHref)){
                        xmlimaHref = "";
                    }
                    res.put("imagesXmlHrefString", xmlimaHref);
                    res.put("content", bodyBlock);
                }else if("publishTime".endsWith(key)&&StringUtils.isNotBlank(val.toString())){//处理时间格式
                    SimpleDateFormat df = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                    Date date = df.parse(val.toString());
                    res.put(key,DateUtil.format(date,"yyyy-MM-dd")) ;
                }else if("rangeName".endsWith(key)){//处理发文范围
                    //获取部门信息
                    NewsPublishRange newsPublishRange = new NewsPublishRange();
                    newsPublishRange.setNewsNoticeId(notice.getId());
                    List<NewsPublishRange> NewsPublishRangeList = newsPublishRangeService.getAll(newsPublishRange);
                    String rangeName = "";
                    for(int i=0;i<NewsPublishRangeList.size();i++){
                        if(i == NewsPublishRangeList.size() - 1){
                            rangeName += NewsPublishRangeList.get(i).getOrgName();
                        }else{
                            rangeName += NewsPublishRangeList.get(i).getOrgName()+";";
                        }
                        res.put(key,rangeName);
                    }
                }else if("approveRemark".endsWith(key)&&StringUtils.isNotBlank(val.toString())){//处理发文审批信息
                    List<CommentVo> approveRecords = new ArrayList<CommentVo>();
                    try {
                        NewsNoticeProcess newsNoticeProcess = new NewsNoticeProcess();
                        newsNoticeProcess.setNewNoticeId(notice.getId());
                        List<NewsNoticeProcess> all = newsNoticeProcessService.getAll(newsNoticeProcess);
                        if (all != null && !all.isEmpty()) {
                            approveRecords=newsNoticeProcessService.getNewsNoticeComments(all.get(0).getTaskId());
                        }

                        String arrpovalTemp = "";
                        if(approveRecords.size()>0){
                            for(CommentVo cv:approveRecords){
                                if(!cv.getType().equals("FLOWEND")){
                                    if(!cv.getType().equals("END")){
                                        if(cv.getTypeName().equals("提交")||cv.getTypeName().equals("审批")){
                                            arrpovalTemp += cv.getUserId() + cv.getTypeName() + "->";
                                        }
                                    }
                                }
                            }
                        }
                        if(StringUtils.isNotBlank(arrpovalTemp)){
                            res.put(key,arrpovalTemp);
                        }else{
                            res.put(key,val);
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }else{
                    if(StringUtils.isNotBlank(val.toString())){
                        val = ExportNoticeUtil.transform(val.toString());
                    }
                    res.put(key,val);
                }

            }
        }catch(Exception e){
            e.printStackTrace();
        }

        return res;
    }

    /**
     * 替换文本中的特殊字符
     * @param str
     * @return
     */
    private static String transform(String str){
        if(str.contains("<")||str.contains(">")||str.contains("&")){
            str=str.replaceAll("&", "&amp;");
            str=str.replaceAll("<", "&lt;");
            str=str.replaceAll(">", "&gt;");
        }
        return str;
    }

    /**
     * 导出Excel
     * @param ids
     * @param response
     * @param newsNoticeService
     * @param newsNoticeProcessService
     * @param iNewsTypeService
     * @param newsFileService
     */
    public static void ExportNoticeExcel(String ids, HttpServletResponse response, INewsNoticeService newsNoticeService, INewsNoticeProcessService newsNoticeProcessService, INewsTypeService iNewsTypeService, INewsFileService newsFileService){

        // 导入2007xecel
        XSSFWorkbook workBook = new XSSFWorkbook();
        try {
            // 生成一个表格
            XSSFSheet xsheet = workBook.createSheet();
            workBook.setSheetName(0, "Sheet1");

            XSSFPrintSetup xhps = xsheet.getPrintSetup();
            xhps.setPaperSize((short) 9); // 设置A4纸

            int i = 0;
            // 设置列宽
            xsheet.setColumnWidth(i++, 2000);
            xsheet.setColumnWidth(i++, 4000);
            xsheet.setColumnWidth(i++, 4000);
            xsheet.setColumnWidth(i++, 4500);
            xsheet.setColumnWidth(i++, 4500);

            xsheet.setColumnWidth(i++, 4000);
            xsheet.setColumnWidth(i++, 4000);
            xsheet.setColumnWidth(i++, 4000);
            xsheet.setColumnWidth(i++, 4000);
            xsheet.setColumnWidth(i++, 6000);

            xsheet.setColumnWidth(i++, 6000);
            xsheet.setColumnWidth(i++, 4000);
            xsheet.setColumnWidth(i++, 5000);
            xsheet.setColumnWidth(i++, 4000);
            xsheet.setColumnWidth(i++, 4000);

            xsheet.setColumnWidth(i++, 4000);
            xsheet.setColumnWidth(i++, 4000);

            // 设置字段字体
            XSSFFont xFieldFont = workBook.createFont();
            xFieldFont.setFontName("微软雅黑");
            xFieldFont.setFontHeightInPoints((short) 11);// 字体大小
            xFieldFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);// 加粗
            // 字段单元格样式
            XSSFCellStyle xFieldStyle = workBook.createCellStyle();
            xFieldStyle.setFont(xFieldFont);
            xFieldStyle.setAlignment(HorizontalAlignment.CENTER);// 左右居中
            xFieldStyle.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
            xFieldStyle.setBorderLeft(BorderStyle.THIN);// 边框的大小
            xFieldStyle.setBorderRight(BorderStyle.THIN);
            xFieldStyle.setBorderTop(BorderStyle.THIN);
            xFieldStyle.setBorderBottom(BorderStyle.THIN);

            // 设置字段字体
            XSSFFont xField2Font = workBook.createFont();
            xField2Font.setFontName("微软雅黑");
            xField2Font.setFontHeightInPoints((short) 10);// 字体大小
            xField2Font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);// 加粗
            // 字段单元格样式
            XSSFCellStyle xField2Style = workBook.createCellStyle();
            xField2Style.setFont(xField2Font);
            xField2Style.setAlignment(HorizontalAlignment.CENTER);// 左右居中
            xField2Style.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
            xField2Style.setBorderLeft(BorderStyle.THIN);// 边框的大小
            xField2Style.setBorderRight(BorderStyle.THIN);
            xField2Style.setBorderTop(BorderStyle.THIN);
            xField2Style.setBorderBottom(BorderStyle.THIN);



            // 设置字段内容字体
            XSSFFont xContentFont = workBook.createFont();
            xContentFont.setFontName("微软雅黑");
            xContentFont.setFontHeightInPoints((short) 10);// 字体大小
            // 字段单元格样式
            XSSFCellStyle xContentStyle = workBook.createCellStyle();
            xContentStyle.setFont(xContentFont);
            xContentStyle.setAlignment(HorizontalAlignment.CENTER);// 左右居中
            xContentStyle.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
            xContentStyle.setWrapText(true);// 自动换行
            xContentStyle.setBorderLeft(BorderStyle.THIN);// 边框的大小
            xContentStyle.setBorderRight(BorderStyle.THIN);
            xContentStyle.setBorderTop(BorderStyle.THIN);
            xContentStyle.setBorderBottom(BorderStyle.THIN);

            XSSFRow xrow1 = null;
            xrow1 = xsheet.createRow(0);
            xrow1.setHeight((short) 400);

            XSSFCell xrow1_cell1 = xrow1.createCell(0);//创建第1行第一列
            xrow1_cell1.setCellValue("亚厦公文发布清单（查询时间："+getTime("yyyy年MM月hh日HH时mm分")+"）");
            xrow1_cell1.setCellStyle(xFieldStyle);

            CellRangeAddress address = new CellRangeAddress(0, 0, 0, i-1);//合并单元格 参数依次为 行start end 列start end
            xsheet.addMergedRegion(address);

            XSSFRow xrow2 = null;
            xrow2 = xsheet.createRow(1);
            xrow2.setHeight((short) 400);

            for(int j=0;j<i;j++){
                XSSFCell xrow2_cell1 = xrow2.createCell(j);
                xrow2_cell1.setCellStyle(xField2Style);
                String temp = "";
                if(j == 0){
                    temp = "序号";
                }else if(j == 1){
                    temp = "发文类型";
                }else if(j == 2){
                    temp = "发文主体";
                }else if(j == 3){
                    temp = "发文编号";
                }else if(j == 4){
                    temp = "标题";
                }else if(j == 5){
                    temp = "发文单位";
                }else if(j == 6){
                    temp = "发文部门";
                }else if(j == 7){
                    temp = "提交人";
                }else if(j == 8){
                    temp = "状态";
                }else if(j == 9){
                    temp = "提交时间";
                }else if(j == 10){
                    temp = "发文时间";
                }else if(j == 11){
                    temp = "发布版块";
                }else if(j == 12){
                    temp = "关键词";
                }else if(j == 13){
                    temp = "有无附件";
                }else if(j == 14){
                    temp = "报送审计时间";
                }else if(j == 15){
                    temp = "报送审计经办人";
                }else if(j == 16){
                    temp = "档号";
                }
                xrow2_cell1.setCellValue(temp);
            }

            String []idArr = ids.split(",");
            for(int k=0;k<idArr.length;k++){
                XSSFRow xrow3 = null;
                xrow3 = xsheet.createRow(k+2);
                NewsNotice notice = newsNoticeService.getNoticeById(idArr[k]);

                for(int j=0;j<i;j++){
                    String temp = "";
                    XSSFCell xrow3_cell1 = xrow3.createCell(j);
                    xrow3_cell1.setCellStyle(xContentStyle);
                    if(j == 0){
                        temp = k+1+"";
                    }else if(j == 1){
                        temp = notice.getCategoryName();
                    }else if(j == 2){
                        temp = notice.getOwnerName();
                    }else if(j == 3){
                        temp = notice.getArticleNo();
                    }else if(j == 4){
                        temp = notice.getTitle();
                    }else if(j == 5){
                        temp = notice.getCreateCompanyName();
                    }else if(j == 6){
                        temp = notice.getCreateDeptmentName();
                    }else if(j == 7){
                        NewsNoticeProcess newsNoticeProcess = new NewsNoticeProcess();
                        newsNoticeProcess.setNewNoticeId(notice.getId());
                        List<NewsNoticeProcess> newsNoticeProcessList = newsNoticeProcessService.getAll(newsNoticeProcess);
                        if(null !=newsNoticeProcessList && newsNoticeProcessList.size()>0){
                            temp = newsNoticeProcessList.get(0).getUsername();
                        }
                    }else if(j == 8){
                        int status = notice.getPublishStatus();
                        if(status == 0) {
                            temp  = "未发布";
                        } else if (status == -1) {
                            temp =  "撤回";
                        } else if (status == -2) {
                            temp =  "终止";
                        } else if (status == 1) {
                            if (notice.getPublishTime().getTime() > new Date().getTime()) {
                                temp =  "待发布";
                            }
                            temp = "已发布";
                        }
                    }else if(j == 9){
                        temp = DateUtil.format( notice.getWritingTime(),"yyyy-MM-dd HH:mm:ss");
                    }else if(j == 10){
                        if(notice.getPublishStatus() == 1){
                            temp = DateUtil.format( notice.getPublishTime(),"yyyy-MM-dd HH:mm:ss");
                        }else{
                            temp = "";
                        }
                    }else if(j == 11){
                        String typeIds = notice.getTypeIdArray();
                        temp = "";
                        if(StringUtils.isNotBlank(typeIds)){
                            String arr[] = typeIds.split(",");
                            for(int arrIndex=0;arrIndex<arr.length;arrIndex++){
                                NewsType type = iNewsTypeService.getNewsTypeById(arr[arrIndex]);
                                if(arrIndex == arr.length-1){
                                    temp += type.getName();
                                }else{
                                    temp += type.getName()+";";
                                }
                            }
                        }
                    }else if(j == 12){
                        temp = notice.getKeyword();
                    }else if(j == 13){
                        NewsFile newsFile = new NewsFile();
                        newsFile.setRefId(notice.getId());
                        PagerModel<NewsFile> pm = null;
                        if (StringUtils.isNotBlank(newsFile.getRefId())) {
                            pm = newsFileService.getPagerModelByQuery(newsFile, new Query());
                        }
                        List<NewsFile> files = pm.getRows();
                        if(null!=files&&files.size()>0){
                            temp = "有";
                        }else{
                            temp = "无";
                        }

                    }else {
                        temp = "";
                    }
                    xrow3_cell1.setCellValue(temp);
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        try {
            String title = "公文清单"+getTime("yyyyMMdd")+".xlsx";
            String formatTitle = StringUtils.isNotBlank(title)?FilePattern.matcher(title).replaceAll("")
                    :"标题为空";
            // 输出Excel文件
            OutputStream output = response.getOutputStream();
            response.reset();
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + new String(formatTitle.getBytes("utf-8"), "ISO8859-1"));
            response.setContentType("application/msexcel");
            workBook.write(output);
            output.close();
            output.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
