package com.viz.jira.app.ppt.service;

import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.jsoup.nodes.Element;

public interface HtmlToPptService {

  void writeHtmlToTableCell(Element element, XSLFTableCell tableCell);
}
