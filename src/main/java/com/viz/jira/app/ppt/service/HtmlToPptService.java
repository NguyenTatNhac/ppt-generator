package com.viz.jira.app.ppt.service;

import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.jsoup.nodes.Document;

public interface HtmlToPptService {

  void writePxtSummary(Document document, XSLFTextShape placeholder);
}
