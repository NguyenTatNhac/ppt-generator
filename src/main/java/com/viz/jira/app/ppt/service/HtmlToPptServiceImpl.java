package com.viz.jira.app.ppt.service;

import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

@Service
public class HtmlToPptServiceImpl implements HtmlToPptService {

  private static final Logger log = LoggerFactory.getLogger(HtmlToPptServiceImpl.class);

  @Override
  public void writeHtmlToTextShape(Element element, XSLFTextShape textShape) {
    log.debug("HTML content to be writing:\n{}", element);
    String tagName = element.tagName();
    switch (tagName) {
      case "p":
        writeParagraphToTextShape(element, textShape);
        break;
      case "ul":
        writeBulletListToTextShape(element, textShape);
        break;
      case "body":
      case "div":
        // If the element is a body or div, it could have multi children.
        writeDivToTextShape(element, textShape);
        break;
      default:
        textShape.clearText();
        log.warn("The HTML tag name [{}] is not yet handled to write in to PPT.", tagName);
    }
  }

  private void writeDivToTextShape(Element element, XSLFTextShape textShape) {
    /* The cell should already have one paragraph with Format set. We edit text of the first
     * paragraph, and append new paragraph after in order to keep the format. */
    Elements children = element.children();

    if (children.isEmpty()) {
      textShape.clearText();
      return;
    }

    for (int i = 0; i < children.size(); i++) {
      Element child = children.get(i);
      if (i == 0) {
        XSLFTextParagraph firstParagraph = textShape.getTextParagraphs().get(0);
        XSLFTextRun firstTextRun = firstParagraph.getTextRuns().get(0);
        firstTextRun.setText(child.text());
      } else {
        textShape.appendText(child.text(), true);
      }
    }
  }

  private void writeBulletListToTextShape(Element ul, XSLFTextShape textShape) {
    /* Each cell has already one bullet point with Format set. We edit text of the first point, and
     * append new point to the cell, in order to keep the format (Font, Font Size,...) */
    Elements children = ul.children();

    if (children.isEmpty()) {
      textShape.clearText();
      return;
    }

    for (int i = 0; i < children.size(); i++) {
      Element li = children.get(i);
      if (i == 0) {
        XSLFTextParagraph firstPoint = textShape.getTextParagraphs().get(0);
        XSLFTextRun firstTextRun = firstPoint.getTextRuns().get(0);
        firstTextRun.setText(li.text());
      } else {
        textShape.appendText(li.text(), true);
      }
    }
  }

  private void writeParagraphToTextShape(Element pElement, XSLFTextShape textShape) {
    setTextKeepFormat(pElement.text(), textShape);
  }

  private void setTextKeepFormat(String text, XSLFTextShape textShape) {
    // Assume the shape has only one paragraph, the paragraph has only one text run
    XSLFTextParagraph paragraph = textShape.getTextParagraphs().get(0);
    XSLFTextRun textRun = paragraph.getTextRuns().get(0);
    textRun.setText(text);
  }
}
