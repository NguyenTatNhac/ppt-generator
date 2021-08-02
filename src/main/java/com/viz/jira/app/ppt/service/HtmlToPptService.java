package com.viz.jira.app.ppt.service;

import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.jsoup.nodes.Document;

public interface HtmlToPptService {

  void writePxtSummary(Document document, XSLFTextShape placeholder);

  void writeCommentBlock(Document document, XSLFTextShape placeholder);

  void writeMilestonesBlock(Document document, XSLFTable table);
}
