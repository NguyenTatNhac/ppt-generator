package com.viz.jira.app.ppt.service;

import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

public interface HtmlToPptService {

  void writeHtmlToTableCell(Element element, XSLFTableCell tableCell);

  void writeCommentBlock(Document document, XSLFTextShape placeholder);

  void writeMilestonesBlock(Document document, XSLFTable table);
}
