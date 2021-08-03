package com.viz.jira.app.ppt.service;

import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;
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
  public void writeCommentBlock(Document document, XSLFTextShape placeholder) {
    log.info("Writing HTML from field [Comment Block] to the PPT slide.");
    log.debug("HTML content to be writing:\n{}", document);
  }

  @Override
  public void writeMilestonesBlock(Document document, XSLFTable table) {
    log.debug("Milestones HTML content to be writing:\n{}", document);
    Elements tables = document.body().getElementsByTag("table");

    if (tables.isEmpty()) {
      log.warn("There is no milestones table found.");
    } else {
      Element htmlTable = tables.first();
    }
  }

  private void writeParagraphToTableCell(Element pElement, XSLFTextShape textShape) {
    XSLFTextParagraph paragraph = textShape.addNewTextParagraph();
    for (Node pNode : pElement.childNodes()) {
      writeNodeToParagraph(pNode, paragraph);
    }
  }

  private void writeNodeToParagraph(Node pNode, XSLFTextParagraph paragraph) {
    if (pNode instanceof TextNode) {
      writeTextNodeToParagraph((TextNode) pNode, paragraph);
    } else if (pNode instanceof Element) {
      writeParagraphElementToParagraph((Element) pNode, paragraph);
    }
  }

  private void writeParagraphElementToParagraph(Element element, XSLFTextParagraph paragraph) {
    // This element can be a <b> or <i>. We only handle the <b> for now and assume it always <b>
    XSLFTextRun textRun = paragraph.addNewTextRun();
    textRun.setText(element.text());

    if (element.tagName().equals("b")) {
      textRun.setBold(true);
    }

    if (element.tagName().equals("i") || element.tagName().equals("em")) {
      textRun.setItalic(true);
    }
  }

  private void writeTextNodeToParagraph(TextNode textNode, XSLFTextParagraph paragraph) {
    XSLFTextRun textRun = paragraph.addNewTextRun();
    textRun.setText(textNode.text());
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
}
