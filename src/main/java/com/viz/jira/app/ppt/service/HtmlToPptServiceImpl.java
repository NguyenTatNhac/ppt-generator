package com.viz.jira.app.ppt.service;

import java.util.List;
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
    /* We expect each TextShape has already one Paragraph with TextRuns. The TextRuns have
     * pre-defined Format (Font, Font Size, Bullet style,...) for reuse. */
    XSLFTextParagraph firstParagraph = textShape.getTextParagraphs().get(0);

    // Make sure that the paragraph has only one text run before writing data
    makeSureOnlyOneTextRunInParagraph(firstParagraph);

    writeHtmlToParagraph(element, firstParagraph);
  }

  private void writeHtmlToParagraph(Element element, XSLFTextParagraph paragraph) {
    String tagName = element.tagName();
    switch (tagName) {
      case "p":
        writeParagraph(element, paragraph);
        break;
      case "ul":
        writeBulletList(element, paragraph);
        break;
      case "body":
      case "div":
        // If the element is a body or div, it could have multi children.
        writeBodyOrDiv(element, paragraph);
        break;
      default:
        paragraph.getTextRuns().clear();
        log.warn("The HTML tag name [{}] is not yet handled to write into PPT.", tagName);
    }
  }

  private void writeBodyOrDiv(Element element, XSLFTextParagraph paragraph) {
    Elements children = element.children();

    if (children.isEmpty()) {
      paragraph.getTextRuns().clear();
      return;
    }

    XSLFTextParagraph currentParagraph;
    XSLFTextParagraph nextParagraph = paragraph;
    for (int i = 0; i < children.size(); i++) {
      Element child = children.get(i);
      currentParagraph = nextParagraph;

      // If this child is not the last one, append a new paragraph to the end
      if (i < children.size() - 1) {
        nextParagraph = paragraph.getParentShape()
            .appendText(" ", true)
            .getParagraph();
      }

      writeHtmlToParagraph(child, currentParagraph);
    }
  }

  private void writeBulletList(Element ul, XSLFTextParagraph paragraph) {
    /* Assume that the paragraph has already one bullet point. We edit text of the first point, and
     * append new point to the cell, in order to keep the format (Font, Font Size,...) */
    Elements children = ul.children();

    if (children.isEmpty()) {
      // Clear all texts in this paragraph
      paragraph.getTextRuns().clear();
      return;
    }

    for (int i = 0; i < children.size(); i++) {
      Element li = children.get(i);
      if (i == 0) {
        XSLFTextRun firstTextRun = paragraph.getTextRuns().get(0);
        firstTextRun.setText(li.text());
      } else {
        paragraph.getParentShape().appendText(li.text(), true);
      }
    }
  }

  private void writeParagraph(Element pElement, XSLFTextParagraph paragraph) {
    /* This paragraph could be pre-defined as the bullet list. Now make it as the normal para. */
    paragraph.setBullet(false);
    setTextKeepFormat(pElement.text(), paragraph);
  }

  private void setTextKeepFormat(String text, XSLFTextParagraph paragraph) {
    // Assume that the paragraph has only one text run
    XSLFTextRun textRun = paragraph.getTextRuns().get(0);
    textRun.setText(text);
  }

  /**
   * Make sure that the paragraph has only one text run by removing others if they are present.
   *
   * @param paragraph The paragraph we need to working on
   */
  private void makeSureOnlyOneTextRunInParagraph(XSLFTextParagraph paragraph) {
    List<XSLFTextRun> textRuns = paragraph.getTextRuns();

    /* Remove all other text runs after index 0 (from the first text run) */
    if (textRuns.size() > 1) {
      textRuns.subList(1, textRuns.size()).clear();
    }
  }
}
