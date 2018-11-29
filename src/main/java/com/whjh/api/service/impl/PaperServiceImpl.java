package com.whjh.api.service.impl;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.TypeReference;
import com.aspose.words.*;
import com.aspose.words.Document;
import com.whjh.api.model.TestLib;
import com.whjh.api.service.PaperService;
import org.apache.commons.lang.StringEscapeUtils;
import org.apache.poi.hdgf.streams.Stream;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.TextNode;
import org.jsoup.select.Elements;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMathPara;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.springframework.stereotype.Service;
import uk.ac.ed.ph.snuggletex.SnuggleEngine;
import uk.ac.ed.ph.snuggletex.SnuggleInput;
import uk.ac.ed.ph.snuggletex.SnuggleSession;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;
import java.awt.*;
import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class PaperServiceImpl implements PaperService {
    private final File stylesheet = new File(PaperServiceImpl.class.getClassLoader().getResource("").getPath() + "MML2OMML.XSL");
    private final TransformerFactory tFactory = TransformerFactory.newInstance();
    private final StreamSource stylesource = new StreamSource(stylesheet);

    /**
     * 生成试卷docx
     * @param jsonObject
     * @return 文件名
     */
    @Override
    public String createPaperDocx(JSONObject jsonObject) {
        // 验证License
        if (!getLicense()) {
            return "";
        }

        try{
            JSONArray jsonArray = jsonObject.getJSONArray("tllist");

            //处理题目列表，根据题型分类
            ArrayList<TestLib> testLibs = JSON.parseObject(jsonArray.toJSONString(), new TypeReference<ArrayList<TestLib>>() {});
            ArrayList<TestLib> choiceTestLibs = new ArrayList<TestLib>();
            ArrayList<TestLib> torfTestLibs = new ArrayList<TestLib>();
            ArrayList<TestLib> fillTestLibs = new ArrayList<TestLib>();
            ArrayList<TestLib> resolveTestLibs = new ArrayList<TestLib>();
            for(int testlibIndex=0;testlibIndex<testLibs.size();testlibIndex++){
                TestLib testLib = testLibs.get(testlibIndex);
                String questiontype = testLib.getQuestiontype();
                if("10".equals(questiontype)){
                    choiceTestLibs.add(testLib);
                }else if("20".equals(questiontype)){
                    choiceTestLibs.add(testLib);
                }else if("30".equals(questiontype)){
                    torfTestLibs.add(testLib);
                }else if("40".equals(questiontype)){
                    fillTestLibs.add(testLib);
                }else if("50".equals(questiontype)){
                    resolveTestLibs.add(testLib);
                }
            }

            HashMap<String,String> scoreTbMap = new HashMap<String,String>();
            scoreTbMap.put("1","一");
            scoreTbMap.put("2","二");
            scoreTbMap.put("3","三");
            scoreTbMap.put("4","四");
            int testlibTypeCount = 0;
            if(choiceTestLibs.size()!=0){
                testlibTypeCount++;
            }
            if(torfTestLibs.size()!=0){
                testlibTypeCount++;
            }
            if(fillTestLibs.size()!=0){
                testlibTypeCount++;
            }
            if(resolveTestLibs.size()!=0){
                testlibTypeCount++;
            }

            String headline = jsonObject.getString("headline");
            String selecthead = jsonObject.getString("selecthead");
            String gapfillinghead = jsonObject.getString("gapfillinghead");
            String freeresponsehead = jsonObject.getString("freeresponsehead");
            String tofhead = jsonObject.getString("tofhead");
            String hassubhead = jsonObject.getString("hassubhead");
            String subhead = jsonObject.getString("subhead");
            String haspaperinfo = jsonObject.getString("haspaperinfo");
            String paperinfo = jsonObject.getString("paperinfo");
            String hasstudentinfo = jsonObject.getString("hasstudentinfo");
            String studentinfo = jsonObject.getString("studentinfo");
            String haspartnote = jsonObject.getString("haspartnote");
            String partonetitle = jsonObject.getString("partonetitle");
            String partonenote = jsonObject.getString("partonenote");
            String parttwotitle = jsonObject.getString("parttwotitle");
            String parttwonote = jsonObject.getString("parttwonote");
            String hastotalscorebar = jsonObject.getString("hastotalscorebar");
            String hasbigqsarea = jsonObject.getString("hasbigqsarea");
            String hasgutter = jsonObject.getString("hasgutter");
            String hascaution = jsonObject.getString("hascaution");
            String caution = jsonObject.getString("caution");
            String hassecrettag = jsonObject.getString("hassecrettag");
            String secrettag = jsonObject.getString("secrettag");
            String showanalysis = jsonObject.getString("showanalysis");
            String showresolve = jsonObject.getString("showresolve");
            String showanswer = jsonObject.getString("showanswer");

            //是否显示答案分析页
            boolean showAnswerPage = true;
            if(!"1".equals(showanalysis) && ! "1".equals(showresolve) && !"1".equals(showanswer)){
                showAnswerPage = false;
            }

            //根据参数选择模板
            String tempFileName = "";
            if("1".equals(hasgutter)){
                if(showAnswerPage){
                    tempFileName = "Template_gutter.docx";
                }else{
                    tempFileName = "Template_gutter_noanswer.docx";
                }
            }else{
                if(showAnswerPage){
                    tempFileName = "Template_nogutter.docx";
                }else{
                    tempFileName = "Template_nogutter_noanswer.docx";
                }

            }
            InputStream is = PaperServiceImpl.class.getClassLoader().getResourceAsStream(tempFileName);
            Document temdoc = new Document(is);
            DocumentBuilder docBuilder = new DocumentBuilder(temdoc);

            //保密标记
            if("1".equals(hassecrettag)){
                docBuilder.moveToBookmark("secrettag");
                docBuilder.write(secrettag);
            }else{
                docBuilder.moveToBookmark("secrettag");
                docBuilder.getCurrentParagraph().remove();
            }

            //大标题
            docBuilder.moveToBookmark("headline");
            docBuilder.write(headline);

            //副标题
            if("1".equals(hassubhead)){
                docBuilder.moveToBookmark("subhead");
                docBuilder.write(subhead);
            }else{
                docBuilder.moveToBookmark("subhead");
                docBuilder.getCurrentParagraph().remove();
            }

            //试卷信息
            if("1".equals(haspaperinfo)){
                docBuilder.moveToBookmark("paperinfo");
                docBuilder.write(paperinfo);
            }else{
                docBuilder.moveToBookmark("paperinfo");
                docBuilder.getCurrentParagraph().remove();
            }

            //考生信息
            if("1".equals(hasstudentinfo)){
                docBuilder.moveToBookmark("studentinfo");
                docBuilder.write(studentinfo);
            }else{
                docBuilder.moveToBookmark("studentinfo");
                docBuilder.getCurrentParagraph().remove();
            }

            //总分栏
            if("1".equals(hastotalscorebar)){
                docBuilder.moveToBookmark("totalscorebar");
                Table table = docBuilder.startTable();

                docBuilder.insertCell();
                table.autoFit(AutoFitBehavior.fromName("FIXED_COLUMN_WIDTHS"));
                table.setAlignment(TableAlignment.CENTER);
                docBuilder.getFont().setBold(true);
                docBuilder.getCellFormat().setWidth(50.0);
                docBuilder.getCellFormat().getBorders().getTop().setLineWidth(0.1);
                docBuilder.getCellFormat().getBorders().getBottom().setLineWidth(0.1);
                docBuilder.write("题号");
                for(int typeIndex=1;typeIndex<=testlibTypeCount;typeIndex++){
                    docBuilder.insertCell();
                    docBuilder.getFont().setBold(true);
                    docBuilder.getCellFormat().setWidth(50.0);
                    docBuilder.write(scoreTbMap.get(new Integer(typeIndex).toString()));
                }
                docBuilder.insertCell();
                docBuilder.getFont().setBold(true);
                docBuilder.getCellFormat().setWidth(50.0);
                docBuilder.write("总分");
                docBuilder.endRow();

                docBuilder.insertCell();
                docBuilder.getFont().setBold(true);
                docBuilder.getCellFormat().setWidth(50.0);
                docBuilder.write("得分");
                for(int typeIndex=1;typeIndex<=testlibTypeCount;typeIndex++){
                    docBuilder.insertCell();
                    docBuilder.getFont().setBold(true);
                    docBuilder.getCellFormat().setWidth(50.0);
                }
                docBuilder.insertCell();
                docBuilder.getFont().setBold(false);
                docBuilder.getCellFormat().setWidth(50.0);
                docBuilder.endRow();
                docBuilder.endTable();

                //插入表格后书签会换行，多出一行空白，删除掉此时书签所在空白行
                docBuilder.moveToBookmark("totalscorebar");
                docBuilder.getCurrentParagraph().remove();
            }else{
                docBuilder.moveToBookmark("totalscorebar");
                docBuilder.getCurrentParagraph().remove();
            }

            //注意事项
            if("1".equals(hascaution)){
                docBuilder.moveToBookmark("caution");
                docBuilder.write(caution);
            }else{
                docBuilder.moveToBookmark("caution");
                docBuilder.getCurrentParagraph().remove();
            }

            //卷标题和注释
            if("1".equals(haspartnote)){
                docBuilder.moveToBookmark("partonetitle");
                docBuilder.write(partonetitle);
                docBuilder.moveToBookmark("partonenote");
                docBuilder.write(partonenote);
                docBuilder.moveToBookmark("parttwotitle");
                docBuilder.write(parttwotitle);
                docBuilder.moveToBookmark("parttwonote");
                docBuilder.write(parttwonote);
            }else{
                docBuilder.moveToBookmark("partonetitle");
                docBuilder.getCurrentParagraph().remove();
                docBuilder.moveToBookmark("partonenote");
                docBuilder.getCurrentParagraph().remove();
                docBuilder.moveToBookmark("parttwotitle");
                docBuilder.getCurrentParagraph().remove();
                docBuilder.moveToBookmark("parttwonote");
                docBuilder.getCurrentParagraph().remove();
            }

            //选择题处理
            if(choiceTestLibs.size() > 0){
                if("1".equals(hasbigqsarea)){
                    //评分栏
                    docBuilder.moveToBookmark("hasbigqsarea1");
                    Table areatable1 = docBuilder.startTable();

                    docBuilder.insertCell();
                    areatable1.autoFit(AutoFitBehavior.fromName("FIXED_COLUMN_WIDTHS"));
                    areatable1.setAlignment(TableAlignment.LEFT);
                    docBuilder.getFont().setBold(true);
                    docBuilder.getCellFormat().setWidth(70.0);
                    docBuilder.getCellFormat().getBorders().getTop().setLineWidth(0.1);
                    docBuilder.getCellFormat().getBorders().getBottom().setLineWidth(0.1);
                    docBuilder.write("评卷人");
                    docBuilder.insertCell();
                    docBuilder.getCellFormat().setWidth(70.0);
                    docBuilder.write("得分");
                    docBuilder.endRow();

                    docBuilder.insertCell();
                    docBuilder.getCellFormat().setWidth(70.0);
                    docBuilder.insertCell();
                    docBuilder.getCellFormat().setWidth(70.0);
                    docBuilder.endRow();
                    docBuilder.endTable();

                    docBuilder.moveToBookmark("hasbigqsarea1");
                    docBuilder.getCurrentParagraph().remove();
                }else{
                    docBuilder.moveToBookmark("hasbigqsarea1");
                    docBuilder.getCurrentParagraph().remove();
                }
                docBuilder.moveToBookmark("selecthead");
                docBuilder.write(selecthead);

                docBuilder.moveToBookmark("selectqlist");
                Iterator<TestLib> choiceIt = choiceTestLibs.iterator();
                String[] optpreArr = {"A．","B. ","C. ","D. "};
                while(choiceIt.hasNext()){
                    TestLib choiceTestLib = choiceIt.next();

                    //题干处理
                    String questionstem = choiceTestLib.getQuestionstem();
                    String sortnum = choiceTestLib.getSortnum();
                    org.jsoup.nodes.Document stemdoc = Jsoup.parse(questionstem);
                    Element rootEle = stemdoc.body().child(0);
                    rootEle.html(sortnum + "．" + rootEle.html());
                    Elements latexImgs = stemdoc.select("img.kfformula");
                    for(int i=0;i<latexImgs.size();i++){
                        Element latexImg = latexImgs.get(i);
                        String latex = latexImg.attr("data-latex");
                        TextNode tn = new TextNode("$" + latex + "$","");
                        latexImg.replaceWith(tn);
                    }
                    docBuilder.getFont().setName("宋体");
                    docBuilder.insertHtml(stemdoc.body().html(),true);

                    //选项处理
                    String choiceoptinfo = choiceTestLib.getChoiceoptinfo();
                    JSONArray optArray = JSON.parseArray(choiceoptinfo);
                    Table optTable = docBuilder.startTable();
                    for(int optIndex=0;optIndex<optArray.size();optIndex++){
                        String opthtml = optArray.getString(optIndex);
                        org.jsoup.nodes.Document optdoc = Jsoup.parse(opthtml);
                        Elements optlatexImgs = optdoc.select("img.kfformula");
                        for(int i=0;i<optlatexImgs.size();i++){
                            Element latexImg = optlatexImgs.get(i);
                            String latex = latexImg.attr("data-latex");
                            TextNode tn = new TextNode("$" + latex + "$","");
                            latexImg.replaceWith(tn);
                        }
                        opthtml = optpreArr[optIndex] + optdoc.body().select("p").html();

                        switch (optIndex){
                            case 0:
                                docBuilder.insertCell();
                                optTable.autoFit(AutoFitBehavior.fromName("FIXED_COLUMN_WIDTHS"));
                                optTable.setAlignment(TableAlignment.LEFT);
                                optTable.setBorders(LineStyle.NONE,0.0, Color.white);
                                docBuilder.getFont().setBold(false);
                                docBuilder.getCellFormat().setWidth(200.0);
                                docBuilder.getCellFormat().setWrapText(true);
                                docBuilder.insertHtml(opthtml,true);
                                break;
                            case 1:
                                docBuilder.insertCell();
                                docBuilder.getCellFormat().setWidth(200.0);
                                docBuilder.insertHtml(opthtml,true);
                                docBuilder.endRow();
                                break;
                            case 2:
                                docBuilder.insertCell();
                                optTable.autoFit(AutoFitBehavior.fromName("FIXED_COLUMN_WIDTHS"));
                                optTable.setAlignment(TableAlignment.LEFT);
                                optTable.setBorders(LineStyle.NONE,0.0, Color.white);
                                docBuilder.getCellFormat().setWidth(200.0);
                                docBuilder.insertHtml(opthtml,true);
                                break;
                            case 3:
                                docBuilder.insertCell();
                                docBuilder.getCellFormat().setWidth(200.0);
                                docBuilder.insertHtml(opthtml,true);
                                docBuilder.endRow();
                                break;
                            case 4:
                                docBuilder.insertCell();
                                optTable.autoFit(AutoFitBehavior.fromName("FIXED_COLUMN_WIDTHS"));
                                optTable.setAlignment(TableAlignment.LEFT);
                                optTable.setBorders(LineStyle.NONE,0.0, Color.white);
                                docBuilder.getCellFormat().setWidth(200.0);
                                docBuilder.insertHtml(opthtml,true);
                                docBuilder.endRow();
                                break;
                        }
                    }
                    docBuilder.endTable();
                }
            }

            //判断题处理
            if(torfTestLibs.size() > 0){
                if("1".equals(hasbigqsarea)){
                    //评分栏
                    docBuilder.moveToBookmark("hasbigqsarea2");
                    Table areatable2 = docBuilder.startTable();

                    docBuilder.insertCell();
                    areatable2.autoFit(AutoFitBehavior.fromName("FIXED_COLUMN_WIDTHS"));
                    areatable2.setAlignment(TableAlignment.LEFT);
                    docBuilder.getFont().setBold(true);
                    docBuilder.getCellFormat().setWidth(70.0);
                    docBuilder.getCellFormat().getBorders().getTop().setLineWidth(0.1);
                    docBuilder.getCellFormat().getBorders().getBottom().setLineWidth(0.1);
                    docBuilder.write("评卷人");
                    docBuilder.insertCell();
                    docBuilder.getCellFormat().setWidth(70.0);
                    docBuilder.write("得分");
                    docBuilder.endRow();

                    docBuilder.insertCell();
                    docBuilder.getCellFormat().setWidth(70.0);
                    docBuilder.insertCell();
                    docBuilder.getCellFormat().setWidth(70.0);
                    docBuilder.endRow();
                    docBuilder.endTable();

                    docBuilder.moveToBookmark("hasbigqsarea2");
                    docBuilder.getCurrentParagraph().remove();
                }else{
                    docBuilder.moveToBookmark("hasbigqsarea2");
                    docBuilder.getCurrentParagraph().remove();
                }
                docBuilder.moveToBookmark("tofhead");
                docBuilder.write(tofhead);

                docBuilder.moveToBookmark("tofqlist");
                Iterator<TestLib> torfIt = torfTestLibs.iterator();
                while(torfIt.hasNext()){
                    TestLib torfTestLib = torfIt.next();
                    //题干处理
                    String questionstem = torfTestLib.getQuestionstem();
                    String sortnum = torfTestLib.getSortnum();
                    org.jsoup.nodes.Document stemdoc = Jsoup.parse(questionstem);
                    Element rootEle = stemdoc.body().child(0);
                    rootEle.html(sortnum + "．" + rootEle.html());
                    Elements latexImgs = stemdoc.select("img.kfformula");
                    for(int i=0;i<latexImgs.size();i++){
                        Element latexImg = latexImgs.get(i);
                        String latex = latexImg.attr("data-latex");
                        TextNode tn = new TextNode("$" + latex + "$","");
                        latexImg.replaceWith(tn);
                    }
                    docBuilder.getFont().setName("宋体");
                    docBuilder.insertHtml(stemdoc.body().html(),true);
                }
            }

            //填空题处理
            if(fillTestLibs.size() > 0){
                if("1".equals(hasbigqsarea)){
                    //评分栏
                    docBuilder.moveToBookmark("hasbigqsarea3");
                    Table areatable2 = docBuilder.startTable();

                    docBuilder.insertCell();
                    areatable2.autoFit(AutoFitBehavior.fromName("FIXED_COLUMN_WIDTHS"));
                    areatable2.setAlignment(TableAlignment.LEFT);
                    docBuilder.getFont().setBold(true);
                    docBuilder.getCellFormat().setWidth(70.0);
                    docBuilder.getCellFormat().getBorders().getTop().setLineWidth(0.1);
                    docBuilder.getCellFormat().getBorders().getBottom().setLineWidth(0.1);
                    docBuilder.write("评卷人");
                    docBuilder.insertCell();
                    docBuilder.getCellFormat().setWidth(70.0);
                    docBuilder.write("得分");
                    docBuilder.endRow();

                    docBuilder.insertCell();
                    docBuilder.getCellFormat().setWidth(70.0);
                    docBuilder.insertCell();
                    docBuilder.getCellFormat().setWidth(70.0);
                    docBuilder.endRow();
                    docBuilder.endTable();

                    docBuilder.moveToBookmark("hasbigqsarea3");
                    docBuilder.getCurrentParagraph().remove();
                }else{
                    docBuilder.moveToBookmark("hasbigqsarea3");
                    docBuilder.getCurrentParagraph().remove();
                }
                docBuilder.moveToBookmark("gapfillinghead");
                docBuilder.write(gapfillinghead);

                docBuilder.moveToBookmark("gapfillingqlist");
                Iterator<TestLib> fillIt = fillTestLibs.iterator();
                while(fillIt.hasNext()){
                    TestLib fillTestLib = fillIt.next();
                    //题干处理
                    String questionstem = fillTestLib.getQuestionstem();
                    String sortnum = fillTestLib.getSortnum();
                    questionstem = questionstem.replaceAll("\\{\\*{3}\\}","__________");
                    org.jsoup.nodes.Document stemdoc = Jsoup.parse(questionstem);
                    Element rootEle = stemdoc.body().child(0);
                    rootEle.html(sortnum + "．" + rootEle.html());
                    Elements latexImgs = stemdoc.select("img.kfformula");
                    for(int i=0;i<latexImgs.size();i++){
                        Element latexImg = latexImgs.get(i);
                        String latex = latexImg.attr("data-latex");
                        TextNode tn = new TextNode("$" + latex + "$","");
                        latexImg.replaceWith(tn);
                    }
                    docBuilder.getFont().setName("宋体");
                    docBuilder.insertHtml(stemdoc.body().html(),true);
                }
            }

            //解答题处理
            if(resolveTestLibs.size() > 0){
                if("1".equals(hasbigqsarea)){
                    //评分栏
                    docBuilder.moveToBookmark("hasbigqsarea4");
                    Table areatable3 = docBuilder.startTable();

                    docBuilder.insertCell();
                    areatable3.autoFit(AutoFitBehavior.fromName("FIXED_COLUMN_WIDTHS"));
                    areatable3.setAlignment(TableAlignment.LEFT);
                    docBuilder.getFont().setBold(true);
                    docBuilder.getCellFormat().setWidth(70.0);
                    docBuilder.getCellFormat().getBorders().getTop().setLineWidth(0.1);
                    docBuilder.getCellFormat().getBorders().getBottom().setLineWidth(0.1);
                    docBuilder.write("评卷人");
                    docBuilder.insertCell();
                    docBuilder.getCellFormat().setWidth(70.0);
                    docBuilder.write("得分");
                    docBuilder.endRow();

                    docBuilder.insertCell();
                    docBuilder.getCellFormat().setWidth(70.0);
                    docBuilder.insertCell();
                    docBuilder.getCellFormat().setWidth(70.0);
                    docBuilder.endRow();
                    docBuilder.endTable();

                    docBuilder.moveToBookmark("hasbigqsarea4");
                    docBuilder.getCurrentParagraph().remove();
                }else{
                    docBuilder.moveToBookmark("hasbigqsarea4");
                    docBuilder.getCurrentParagraph().remove();
                }
                docBuilder.moveToBookmark("freeresponsehead");
                docBuilder.write(freeresponsehead);

                docBuilder.moveToBookmark("freeresponseqlist");
                Iterator<TestLib> resolveIt = resolveTestLibs.iterator();
                while(resolveIt.hasNext()){
                    TestLib resolveTestLib = resolveIt.next();
                    //题干处理
                    String questionstem = resolveTestLib.getQuestionstem();
                    String sortnum = resolveTestLib.getSortnum();
                    org.jsoup.nodes.Document stemdoc = Jsoup.parse(questionstem);
                    Element rootEle = stemdoc.body().child(0);
                    rootEle.html(sortnum + "．" + rootEle.html());
                    Elements latexImgs = stemdoc.select("img.kfformula");
                    for(int i=0;i<latexImgs.size();i++){
                        Element latexImg = latexImgs.get(i);
                        String latex = latexImg.attr("data-latex");
                        TextNode tn = new TextNode("$" + latex + "$","");
                        latexImg.replaceWith(tn);
                    }
                    docBuilder.getFont().setName("宋体");
                    docBuilder.insertHtml(stemdoc.body().html(),true);
                }
            }

            if(showAnswerPage){
                //答案解析页大标题
                docBuilder.moveToBookmark("answerheadline");
                docBuilder.write(headline);

                if(choiceTestLibs.size() > 0){
                    docBuilder.moveToBookmark("answerselecthead");
                    docBuilder.write(selecthead);

                    docBuilder.moveToBookmark("answerselectqlist");
                    Iterator<TestLib> choiceIt = choiceTestLibs.iterator();
                    String[] optpreArr = {"A. ","B. ","C. ","D. "};
                    while(choiceIt.hasNext()){
                        TestLib choiceTestLib = choiceIt.next();

                        //题干处理
                        String questionstem = choiceTestLib.getQuestionstem();
                        String sortnum = choiceTestLib.getSortnum();
                        org.jsoup.nodes.Document stemdoc = Jsoup.parse(questionstem);
                        Element rootEle = stemdoc.body().child(0);
                        rootEle.html(sortnum + "．" + rootEle.html());
                        Elements latexImgs = stemdoc.select("img.kfformula");
                        for(int i=0;i<latexImgs.size();i++){
                            Element latexImg = latexImgs.get(i);
                            String latex = latexImg.attr("data-latex");
                            TextNode tn = new TextNode("$" + latex + "$","");
                            latexImg.replaceWith(tn);
                        }
                        docBuilder.getFont().setName("宋体");
                        docBuilder.insertHtml(stemdoc.body().html(),true);

                        //选项处理
                        String choiceoptinfo = choiceTestLib.getChoiceoptinfo();
                        JSONArray optArray = JSON.parseArray(choiceoptinfo);
                        Table optTable = docBuilder.startTable();
                        for(int optIndex=0;optIndex<optArray.size();optIndex++){
                            String opthtml = optArray.getString(optIndex);
                            org.jsoup.nodes.Document optdoc = Jsoup.parse(opthtml);
                            Elements optlatexImgs = optdoc.select("img.kfformula");
                            for(int i=0;i<optlatexImgs.size();i++){
                                Element latexImg = optlatexImgs.get(i);
                                String latex = latexImg.attr("data-latex");
                                TextNode tn = new TextNode("$" + latex + "$","");
                                latexImg.replaceWith(tn);
                            }
                            opthtml = optpreArr[optIndex] + optdoc.body().select("p").html();

                            switch (optIndex){
                                case 0:
                                    docBuilder.insertCell();
                                    optTable.autoFit(AutoFitBehavior.fromName("FIXED_COLUMN_WIDTHS"));
                                    optTable.setAlignment(TableAlignment.LEFT);
                                    optTable.setBorders(LineStyle.NONE,0.0, Color.white);
                                    docBuilder.getFont().setBold(false);
                                    docBuilder.getCellFormat().setWidth(200.0);
                                    docBuilder.getCellFormat().setWrapText(true);
                                    docBuilder.insertHtml(opthtml,true);
                                    break;
                                case 1:
                                    docBuilder.insertCell();
                                    docBuilder.getCellFormat().setWidth(200.0);
                                    docBuilder.insertHtml(opthtml,true);
                                    docBuilder.endRow();
                                    break;
                                case 2:
                                    docBuilder.insertCell();
                                    optTable.autoFit(AutoFitBehavior.fromName("FIXED_COLUMN_WIDTHS"));
                                    optTable.setAlignment(TableAlignment.LEFT);
                                    optTable.setBorders(LineStyle.NONE,0.0, Color.white);
                                    docBuilder.getCellFormat().setWidth(200.0);
                                    docBuilder.insertHtml(opthtml,true);
                                    break;
                                case 3:
                                    docBuilder.insertCell();
                                    docBuilder.getCellFormat().setWidth(200.0);
                                    docBuilder.insertHtml(opthtml,true);
                                    docBuilder.endRow();
                                    break;
                                case 4:
                                    docBuilder.insertCell();
                                    optTable.autoFit(AutoFitBehavior.fromName("FIXED_COLUMN_WIDTHS"));
                                    optTable.setAlignment(TableAlignment.LEFT);
                                    optTable.setBorders(LineStyle.NONE,0.0, Color.white);
                                    docBuilder.getCellFormat().setWidth(200.0);
                                    docBuilder.insertHtml(opthtml,true);
                                    docBuilder.endRow();
                                    break;
                            }
                        }
                        docBuilder.endTable();
                        if("1".equals(showanalysis)){
                            String analysis = choiceTestLib.getAnalysis();
                            org.jsoup.nodes.Document analysisdoc = Jsoup.parse(analysis);
                            analysis = "<p><font color='blue'>【分析】</font>" + analysisdoc.body().select("p").html() + "</p>";
                            docBuilder.insertHtml(analysis,true);
                        }
                        if("1".equals(showresolve)){
                            String resolve = choiceTestLib.getResolve();
                            org.jsoup.nodes.Document resolvedoc = Jsoup.parse(resolve);
                            resolve = "<p><font color='blue'>【解答】</font>" + resolvedoc.body().select("p").html() + "</p>";;
                            docBuilder.insertHtml(resolve,true);
                        }
                        if("1".equals(showanswer)){
                            String answer = choiceTestLib.getChoiceanswer();
                            answer = "<p><font color='blue'>【答案】</font>" + answer + "</p>";;
                            docBuilder.insertHtml(answer,true);
                        }
                        docBuilder.insertHtml("<br>",true);
                    }
                }

                if(torfTestLibs.size() > 0){
                    docBuilder.moveToBookmark("answertofhead");
                    docBuilder.write(tofhead);

                    docBuilder.moveToBookmark("answertofqlist");
                    Iterator<TestLib> torfIt = torfTestLibs.iterator();
                    while(torfIt.hasNext()){
                        TestLib torfTestLib = torfIt.next();
                        //题干处理
                        String questionstem = torfTestLib.getQuestionstem();
                        String sortnum = torfTestLib.getSortnum();
                        org.jsoup.nodes.Document stemdoc = Jsoup.parse(questionstem);
                        Element rootEle = stemdoc.body().child(0);
                        rootEle.html(sortnum + "．" + rootEle.html());
                        Elements latexImgs = stemdoc.select("img.kfformula");
                        for(int i=0;i<latexImgs.size();i++){
                            Element latexImg = latexImgs.get(i);
                            String latex = latexImg.attr("data-latex");
                            TextNode tn = new TextNode("$" + latex + "$","");
                            latexImg.replaceWith(tn);
                        }
                        docBuilder.getFont().setName("宋体");
                        docBuilder.insertHtml(stemdoc.body().html(),true);

                        if("1".equals(showanalysis)){
                            String analysis = torfTestLib.getAnalysis();
                            org.jsoup.nodes.Document analysisdoc = Jsoup.parse(analysis);
                            analysis = "<p><font color='blue'>【分析】</font>" + analysisdoc.body().select("p").html() + "</p>";
                            docBuilder.insertHtml(analysis,true);
                        }
                        if("1".equals(showresolve)){
                            String resolve = torfTestLib.getResolve();
                            org.jsoup.nodes.Document resolvedoc = Jsoup.parse(resolve);
                            resolve = "<p><font color='blue'>【解答】</font>" + resolvedoc.body().select("p").html() + "</p>";;
                            docBuilder.insertHtml(resolve,true);
                        }
                        if("1".equals(showanswer)){
                            String answer = torfTestLib.getTorfanswer();
                            if("10".equals(answer)){
                                answer = "正确";
                            }else if("90".equals(answer)){
                                answer = "错误";
                            }
                            answer = "<p><font color='blue'>【答案】</font>" + answer + "</p>";;
                            docBuilder.insertHtml(answer,true);
                        }
                        docBuilder.insertHtml("<br>",true);
                    }
                }

                if(fillTestLibs.size() > 0){
                    docBuilder.moveToBookmark("answergapfillinghead");
                    docBuilder.write(gapfillinghead);

                    docBuilder.moveToBookmark("answergapfillingqlist");
                    Iterator<TestLib> fillIt = fillTestLibs.iterator();
                    while(fillIt.hasNext()){
                        TestLib fillTestLib = fillIt.next();
                        //题干处理
                        String questionstem = fillTestLib.getQuestionstem();
                        String sortnum = fillTestLib.getSortnum();
                        questionstem = questionstem.replaceAll("\\{\\*{3}\\}","__________");
                        org.jsoup.nodes.Document stemdoc = Jsoup.parse(questionstem);
                        Element rootEle = stemdoc.body().child(0);
                        rootEle.html(sortnum + "．" + rootEle.html());
                        Elements latexImgs = stemdoc.select("img.kfformula");
                        for(int i=0;i<latexImgs.size();i++){
                            Element latexImg = latexImgs.get(i);
                            String latex = latexImg.attr("data-latex");
                            TextNode tn = new TextNode("$" + latex + "$","");
                            latexImg.replaceWith(tn);
                        }
                        docBuilder.getFont().setName("宋体");
                        docBuilder.insertHtml(stemdoc.body().html(),true);

                        if("1".equals(showanalysis)){
                            String analysis = fillTestLib.getAnalysis();
                            org.jsoup.nodes.Document analysisdoc = Jsoup.parse(analysis);
                            analysis = "<p><font color='blue'>【分析】</font>" + analysisdoc.body().select("p").html() + "</p>";
                            docBuilder.insertHtml(analysis,true);
                        }
                        if("1".equals(showresolve)){
                            String resolve = fillTestLib.getResolve();
                            org.jsoup.nodes.Document resolvedoc = Jsoup.parse(resolve);
                            resolve = "<p><font color='blue'>【解答】</font>" + resolvedoc.body().select("p").html() + "</p>";;
                            docBuilder.insertHtml(resolve,true);
                        }
                        if("1".equals(showanswer)){
                            String answer = fillTestLib.getFillanswer();
                            JSONArray answerArr = JSON.parseArray(answer);
                            String strArr = "";
                            for(int answerIndex=0;answerIndex<answerArr.size();answerIndex++){
                                strArr += answerArr.get(answerIndex);
                                if(answerIndex!=answerArr.size()-1){
                                    strArr += "，";
                                }
                            }
                            answer = "<p><font color='blue'>【答案】</font>" + strArr + "</p>";;
                            docBuilder.insertHtml(answer,true);
                        }
                        docBuilder.insertHtml("<br>",true);
                    }
                }

                if(resolveTestLibs.size() > 0){
                    docBuilder.moveToBookmark("answerfreeresponsehead");
                    docBuilder.write(freeresponsehead);

                    docBuilder.moveToBookmark("answerfreeresponseqlist");
                    Iterator<TestLib> resolveIt = resolveTestLibs.iterator();
                    while(resolveIt.hasNext()){
                        TestLib resolveTestLib = resolveIt.next();
                        //题干处理
                        String questionstem = resolveTestLib.getQuestionstem();
                        String sortnum = resolveTestLib.getSortnum();
                        org.jsoup.nodes.Document stemdoc = Jsoup.parse(questionstem);
                        Element rootEle = stemdoc.body().child(0);
                        rootEle.html(sortnum + "．" + rootEle.html());
                        Elements latexImgs = stemdoc.select("img.kfformula");
                        for(int i=0;i<latexImgs.size();i++){
                            Element latexImg = latexImgs.get(i);
                            String latex = latexImg.attr("data-latex");
                            TextNode tn = new TextNode("$" + latex + "$","");
                            latexImg.replaceWith(tn);
                        }
                        docBuilder.getFont().setName("宋体");
                        docBuilder.insertHtml(stemdoc.body().html(),true);

                        if("1".equals(showanalysis)){
                            String analysis = resolveTestLib.getAnalysis();
                            org.jsoup.nodes.Document analysisdoc = Jsoup.parse(analysis);
                            analysis = "<p><font color='blue'>【分析】</font>" + analysisdoc.body().select("p").html() + "</p>";
                            docBuilder.insertHtml(analysis,true);
                        }
                        if("1".equals(showresolve)){
                            String resolve = resolveTestLib.getResolve();
                            org.jsoup.nodes.Document resolvedoc = Jsoup.parse(resolve);
                            resolve = "<p><font color='blue'>【解答】</font>" + resolvedoc.body().select("p").html() + "</p>";;
                            docBuilder.insertHtml(resolve,true);
                        }
                        docBuilder.insertHtml("<br>",true);
                    }
                }
            }

            String filename = PaperServiceImpl.class.getClassLoader().getResource("").getPath() + "PaperWord" + new Date().getTime() + ".docx";
            temdoc.save(filename);

            //对生成的文档中的latex公式进行转换
            FileInputStream in = new FileInputStream(new File(filename));
            XWPFDocument xwpfDocument = new XWPFDocument(in);
            OutputStream os = new FileOutputStream(new File(filename));
            Iterator<XWPFParagraph> iterator = xwpfDocument.getParagraphsIterator();
            convertLatex(iterator);

            Iterator<XWPFTable> iterator2 = xwpfDocument.getTablesIterator();
            XWPFTable tb;
            while (iterator2.hasNext()) {
                tb = iterator2.next();
                Matcher matcher;

                if (latexMatcher(tb.getText()).find()) {
                    java.util.List<XWPFTableRow> rows = tb.getRows();
                    for(int rowIndex=0;rowIndex<rows.size();rowIndex++){
                        XWPFTableRow row = rows.get(rowIndex);
                        java.util.List<XWPFTableCell> cells = row.getTableCells();
                        for(int cellIndex=0;cellIndex<cells.size();cellIndex++){
                            XWPFTableCell cell = cells.get(cellIndex);
                            if (latexMatcher(cell.getText()).find()) {
//                                System.out.println("celltext:" + cell.getText());
                                java.util.List<XWPFParagraph> paragraphList = cell.getParagraphs();
                                Iterator<XWPFParagraph> paragraphIterator = paragraphList.iterator();
                                convertLatex(paragraphIterator);
                            }
                        }
                    }
                }
            }


            xwpfDocument.write(os);
            os.flush();
            os.close();

            return filename;
//            }
        }catch (Exception e){
            e.printStackTrace();
            return "";
        }
    }

    /**
     * 获取license
     *
     * @return
     */
    public boolean getLicense() {
        boolean result = false;
        try {
            InputStream is = PaperServiceImpl.class.getClassLoader().getResourceAsStream("license.xml");
            License aposeLic = new License();
            aposeLic.setLicense(is);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    public Matcher latexMatcher(String str) {
        Pattern pattern = Pattern.compile("\\$(.+?)\\$", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }

    /**latex转成MathML**/
    public String latexToMathMl(String latex) throws Exception {
        //String latex = "$\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}$";
        String mathMl = null;

//        latex = latex.replace(" ","");
//        System.out.println(latex);
        /* Create vanilla SnuggleEngine and new SnuggleSession */
        SnuggleEngine engine = new SnuggleEngine();
        SnuggleSession session = engine.createSession();

        /* Parse some very basic Math Mode input */
        SnuggleInput input = new SnuggleInput(latex);
        session.parseInput(input);

        /* Convert the results to an XML String, which in this case will
         * be a single MathML <math>...</math> element. */
        mathMl = session.buildXMLString();
//        mathMl = fmath.conversion.ConvertFromLatexToMathML.convertToMathML(latex);
//        mathMl = mathMl.replaceFirst("<math ", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" ");
//        mathMl = StringEscapeUtils.unescapeHtml(mathMl);
//        System.out.println(mathMl);
//        mathMl = mathMl.replaceAll("&plusmn;","±");
////        mathMl = mathMl.replaceAll("&sum;","∑");
////        mathMl = mathMl.replaceAll("&int;","∫");
        return mathMl;
    }

    /**将mathml转为word的ooxml格式**/
    public String mathMLToOOXML(String mathMl){
        try{
            Transformer transformer = tFactory.newTransformer(stylesource);
            StringReader stringreader = new StringReader(mathMl);
            StreamSource source = new StreamSource(stringreader);

            StringWriter stringwriter = new StringWriter();
            StreamResult result = new StreamResult(stringwriter);

            transformer.transform(source, result);

            String ooML = stringwriter.toString();
            stringwriter.close();
            return ooML;
        }catch (Exception e){
            e.printStackTrace();
            return "";
        }
    }

    public void createMathDocNode(String runText, CTP ctp) throws Exception {
        Matcher matcher = latexMatcher(runText);
        if (matcher.find()) {
            String latex = matcher.group();
            String mathMl = latexToMathMl(latex);
            String ooxmlMath = mathMLToOOXML(mathMl);

            org.dom4j.Document doc = org.dom4j.DocumentHelper.parseText(ooxmlMath);
            java.util.List nary_e_list = doc.selectNodes("//m:nary/m:e");
            Iterator nary_e_iterator = nary_e_list.iterator();
            while(nary_e_iterator.hasNext()){
                org.dom4j.Element ele = (org.dom4j.Element) nary_e_iterator.next();
                if("".equals(ele.getStringValue())){
                    org.dom4j.Element r_ele = ele.addElement("m:r");
                    r_ele.addElement("m:t").setText(" ");
                }
            }

            java.util.List subsup_e_list = doc.selectNodes("//m:sSubSup/m:e");
            Iterator subsup_e_iterator = subsup_e_list.iterator();
            while(subsup_e_iterator.hasNext()){
                org.dom4j.Element ele = (org.dom4j.Element) subsup_e_iterator.next();
                if("".equals(ele.getStringValue())){
                    org.dom4j.Element r_ele = ele.addElement("m:r");
                    r_ele.addElement("m:t").setText(" ");
                }
            }

            ooxmlMath = doc.asXML();

            int start = matcher.start();
            if(start != 0){
                String prefix = runText.substring(0,start);
                org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR prefixctr = ctp.addNewR();
                prefixctr.addNewRPr().addNewRFonts().setAscii("宋体");
                CTText cttext_pre = prefixctr.addNewT();
                cttext_pre.setStringValue(prefix);
            }

            CTOMathPara ctOMathPara = CTOMathPara.Factory.parse(ooxmlMath);
            CTOMath ctOMath = ctOMathPara.getOMathArray(0);
            XmlCursor xmlcursor = ctOMath.newCursor();
            while (xmlcursor.hasNextToken()) {
                XmlCursor.TokenType tokentype = xmlcursor.toNextToken();
                if (tokentype.isStart()) {
                    if (xmlcursor.getObject() instanceof CTR) {
                        CTR cTR = (CTR) xmlcursor.getObject();
                        cTR.addNewRPr2().addNewRFonts().setAscii("Cambria Math");
                        cTR.getRPr2().getRFonts().setHAnsi("Cambria Math");
                    }
                }
            }
            System.out.println(ctOMath.toString());
            CTOMath ctoMath = ctp.addNewOMath();
            ctoMath.set(ctOMath);

            int end = matcher.end();
            if(end != runText.length()){
                String sufix = runText.substring(end);
                matcher = latexMatcher(sufix);
                if(matcher.find()){
                    createMathDocNode(sufix,ctp);
                }else{
                    org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR sufixctr = ctp.addNewR();
                    sufixctr.addNewRPr().addNewRFonts().setAscii("宋体");
                    CTText cttext_su = sufixctr.addNewT();
                    cttext_su.setStringValue(sufix);
                }
            }
//                        }else{
//                            for(int j=1;j<=groupCount;j++){
//                                String latex = matcher.group(j);
//                                String mathMl = latexToMathMl(latex);
//                                String ooxmlMath = mathMLToOOXML(mathMl);
////                                        CTP ctp = para.getCTP();
//                                if(j == 1){
//                                    int start = matcher.start(j);
//                                    if(start != 0){
//                                        String prefix = runText.substring(0,start);
//                                        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR prectr = ctp.addNewR();
//                                        prectr.addNewRPr().addNewRFonts().setAscii("宋体");
//                                        CTText cttext1 = prectr.addNewT();
//                                        cttext1.setStringValue(prefix);
//                                    }
//                                }
//                                CTOMathPara ctOMathPara = CTOMathPara.Factory.parse(ooxmlMath);
//                                CTOMath ctOMath = ctOMathPara.getOMathArray(0);
//                                XmlCursor xmlcursor = ctOMath.newCursor();
//                                while (xmlcursor.hasNextToken()) {
//                                    XmlCursor.TokenType tokentype = xmlcursor.toNextToken();
//                                    if (tokentype.isStart()) {
//                                        if (xmlcursor.getObject() instanceof CTR) {
//                                            CTR cTR = (CTR) xmlcursor.getObject();
//                                            cTR.addNewRPr2().addNewRFonts().setAscii("Cambria Math");
//                                            cTR.getRPr2().getRFonts().setHAnsi("Cambria Math");
//                                        }
//                                    }
//                                }
//
//                                if(j == groupCount){
//                                    int end = matcher.end();
//                                    if(end != runText.length()){
//                                        String sufix = runText.substring(end);
//                                        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR sufixctr = ctp.addNewR();
//                                        sufixctr.addNewRPr().addNewRFonts().setAscii("宋体");
//                                        CTText cttext_su = sufixctr.addNewT();
//                                        cttext_su.setStringValue(sufix);
//                                    }
//                                }
//                            }
//                        }
            //重新插入run里内容格式可能与原来模板的格式不一致
//                                para.insertNewRun(i).setText(runText);
        }else{
            org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR prefixctr = ctp.addNewR();
            CTText cttext_pre = prefixctr.addNewT();
            cttext_pre.setStringValue(runText);
        }
    }

    public void convertLatex(Iterator<XWPFParagraph> iterator) throws Exception {
        XWPFParagraph para;
        while (iterator.hasNext()) {
            para = iterator.next();
            Matcher matcher;
            if (latexMatcher(para.getParagraphText()).find()) {
                java.util.List<XWPFRun> runs = para.getRuns();
                int removeCount = 1;
                for (int i=0; i<runs.size(); i++) {
                    XWPFRun run = runs.get(i);
                    String runText = run.toString();

//                    System.out.println(runText);
                    CTP ctp = para.getCTP();
                    createMathDocNode(runText,ctp);
                }
                removeCount = runs.size();
                for(int r=0;r<removeCount;r++){
//                    System.out.println(r + "-before: " + para.getText());
                    para.removeRun(0);
//                    System.out.println(r + "-after: " + para.getText());
                }
            }
        }
    }

    public static void main(String[] args){
//        System.out.println(StringEscapeUtils.unescapeHtml("&sum;"));
        try{
            String latex = "$\\int^{n}_{k}x";

            /* Create vanilla SnuggleEngine and new SnuggleSession */
            SnuggleEngine engine = new SnuggleEngine();
            SnuggleSession session = engine.createSession();

            /* Parse some very basic Math Mode input */
            SnuggleInput input = new SnuggleInput(latex);
            session.parseInput(input);

            /* Convert the results to an XML String, which in this case will
             * be a single MathML <math>...</math> element. */
            String xmlString = session.buildXMLString();
            System.out.println("Input " + input.getString()
                    + " was converted to:\n" + xmlString);
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
