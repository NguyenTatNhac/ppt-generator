package com.viz.jira.app.ppt.service;

import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

@Service
public class HtmlToPptServiceImpl implements HtmlToPptService {

  private static final Logger log = LoggerFactory.getLogger(HtmlToPptServiceImpl.class);

  @Override
  public void writeHtmlToTableCell(Element element, XSLFTableCell tableCell) {
    log.debug("HTML content to be writing:\n{}", element);
    String tagName = element.tagName();
    switch (tagName) {
      case "p":
        writeParagraphToTableCell(element, tableCell);
        break;
      case "ul":
        writeBulletListToTableCell(element, tableCell);
        break;
      default:
        log.warn("The HTML tag name [{}] is not yet handled to write in to PPT.", tagName);
    }
  }

  @Override
  public void writeMilestonesBlock(Document document, XSLFTable table) {
    log.debug("Milestones HTML content to be writing:\n{}", document);
  }

  private void writeBulletListToTableCell(Element ul, XSLFTableCell tableCell) {
    /* Each cell has already one bullet point with Format set. We edit text of the first point, and
     * append new point to the cell, in order to keep the format (Font, Font Size,...) */
    Elements children = ul.children();

    if (children.isEmpty()) {
      tableCell.clearText();
      return;
    }

    for (int i = 0; i < children.size(); i++) {
      Element li = children.get(i);
      if (i == 0) {
        XSLFTextParagraph firstPoint = tableCell.getTextParagraphs().get(0);
        XSLFTextRun firstTextRun = firstPoint.getTextRuns().get(0);
        firstTextRun.setText(li.text());
      } else {
        tableCell.appendText(li.text(), true);
      }
    }
  }

  private void writeParagraphToTableCell(Element pElement, XSLFTableCell tableCell) {
    setTextKeepFormat(pElement.text(), tableCell);
  }

  private void setTextKeepFormat(String text, XSLFTableCell tableCell) {
    // Assume the cell has only one paragraph, the paragraph has only one text run
    XSLFTextParagraph paragraph = tableCell.getTextParagraphs().get(0);
    XSLFTextRun textRun = paragraph.getTextRuns().get(0);
    textRun.setText(text);
  }
}
