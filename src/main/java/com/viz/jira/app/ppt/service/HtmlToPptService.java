package com.viz.jira.app.ppt.service;

import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.jsoup.nodes.Element;

public interface HtmlToPptService {

  void writeHtmlToTextShape(Element element, XSLFTextShape textShape);
}
