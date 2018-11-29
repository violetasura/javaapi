package com.whjh.api.model;

/**
 * @author Lee
 * @date 2018/11/14
 */
public class TestLib {
    /**
     * 题目id
     */
    private String tlid;

    /**
     * 题型
     */
    private String questiontype;

    /**
     * 题目分数
     */
    private String scroe;

    /**
     * 题目排序
     */
    private String sortnum;

    /**
     * 题干
     */
    private String questionstem;

    /**
     * 分析
     */
    private String analysis;

    /**
     * 解答
     */
    private String resolve;

    /**
     * 选择题选项数
     */
    private String choiceoptionnum;

    /**
     * 选择题选项内容（字符串数组）
     */
    private String choiceoptinfo;

    /**
     * 选择题答案
     */
    private String choiceanswer;

    /**
     * 判断题答案
     */
    private String torfanswer;

    /**
     * 填空数
     */
    private String fillnum;

    /**
     * 填空题答案（字符串数组）
     */
    private String fillanswer;

    public String getTlid() {
        return tlid;
    }

    public void setTlid(String tlid) {
        this.tlid = tlid;
    }

    public String getQuestiontype() {
        return questiontype;
    }

    public void setQuestiontype(String questiontype) {
        this.questiontype = questiontype;
    }

    public String getScroe() {
        return scroe;
    }

    public void setScroe(String scroe) {
        this.scroe = scroe;
    }

    public String getSortnum() {
        return sortnum;
    }

    public void setSortnum(String sortnum) {
        this.sortnum = sortnum;
    }

    public String getQuestionstem() {
        return questionstem;
    }

    public void setQuestionstem(String questionstem) {
        this.questionstem = questionstem;
    }

    public String getAnalysis() {
        return analysis;
    }

    public void setAnalysis(String analysis) {
        this.analysis = analysis;
    }

    public String getResolve() {
        return resolve;
    }

    public void setResolve(String resolve) {
        this.resolve = resolve;
    }

    public String getChoiceoptionnum() {
        return choiceoptionnum;
    }

    public void setChoiceoptionnum(String choiceoptionnum) {
        this.choiceoptionnum = choiceoptionnum;
    }

    public String getChoiceoptinfo() {
        return choiceoptinfo;
    }

    public void setChoiceoptinfo(String choiceoptinfo) {
        this.choiceoptinfo = choiceoptinfo;
    }

    public String getChoiceanswer() {
        return choiceanswer;
    }

    public void setChoiceanswer(String choiceanswer) {
        this.choiceanswer = choiceanswer;
    }

    public String getTorfanswer() {
        return torfanswer;
    }

    public void setTorfanswer(String torfanswer) {
        this.torfanswer = torfanswer;
    }

    public String getFillnum() {
        return fillnum;
    }

    public void setFillnum(String fillnum) {
        this.fillnum = fillnum;
    }

    public String getFillanswer() {
        return fillanswer;
    }

    public void setFillanswer(String fillanswer) {
        this.fillanswer = fillanswer;
    }
}
