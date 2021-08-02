package com.viz.jira.app.ppt.service;

import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

@Service
public class HtmlToPptServiceImpl implements HtmlToPptService {

  private static final Double LIST_LEFT_MARGIN = 10D;
  private static final Double LIST_INDENT = -10D;
  private static final Double LIST_SPACE_BEFORE = 27.52; // 27.52 / 6.88 = 4.03 pt

  private static final Logger log = LoggerFactory.getLogger(HtmlToPptServiceImpl.class);

  @Override
  public void writePxtSummary(Document document, XSLFTextShape placeholder) {
    log.info("Writing HTML from field [PXT Summary] to the PPT slide.");
    log.debug("HTML content to be writing:\n{}", document);
    writeHtmlToTextShape(document, placeholder);
  }

  @Override
  public void writeCommentBlock(Document document, XSLFTextShape placeholder) {
    log.info("Writing HTML from field [Comment Block] to the PPT slide.");
    log.debug("HTML content to be writing:\n{}", document);
    writeHtmlToTextShape(document, placeholder);
  }

  private void writeHtmlToTextShape(Document document, XSLFTextShape textShape) {
    // Clear everything in the Text Shape before writing new content
    textShape.clearText();

    // We only write body to the Text Shape
    Element body = document.body();
    for (Element child : body.children()) {
      writeHtmlToTextShape(child, textShape);
    }
  }

  private void writeHtmlToTextShape(Element element, XSLFTextShape textShape) {
    String tagName = element.tagName();
    switch (tagName) {
      case "p":
        writeParagraph(element, textShape);
        break;
      case "ul":
        writeBulletList(element, textShape);
        break;
      default:
        log.warn("The HTML tag name [{}] is not yet handled to write in to PPT.", tagName);
    }
  }

  private void writeParagraph(Element pElement, XSLFTextShape textShape) {
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

  private void writeBulletList(Element ul, XSLFTextShape placeholder) {
    for (Element li : ul.children()) {
      XSLFTextParagraph list = placeholder.addNewTextParagraph();
      list.setBullet(true);
      list.setIndentLevel(0);
      list.setLeftMargin(LIST_LEFT_MARGIN);
      list.setIndent(LIST_INDENT);
      list.setSpaceBefore(LIST_SPACE_BEFORE);
      list.setFontAlign(TextParagraph.FontAlign.TOP);
      list.setTextAlign(TextParagraph.TextAlign.LEFT);
      XSLFTextRun textRun = list.addNewTextRun();
      textRun.setText(li.text());
    }
  }
}
