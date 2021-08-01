package com.viz.jira.app.ppt.service;

import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.jsoup.nodes.Element;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

@Service
public class HtmlToPptServiceImpl implements HtmlToPptService {

  private static final Double LIST_LEFT_MARGIN = 10D;
  private static final Double LIST_INDENT = -10D;
  private static final Double LIST_SPACE_BEFORE = 5D;

  private static final Logger log = LoggerFactory.getLogger(HtmlToPptServiceImpl.class);

  /**
   * Write the HTML of field PXT Summary to the given placeholder in the PPTX file. The HTML is in
   * the below structure:
   * <body>
   * <p><b>Header 1</b></p>
   * <ul>
   *   <li>Point 1</li>
   *   <li>Point 2</li>
   * </ul>
   *
   * <p><b>Header 2</b></p>
   * <ul>
   *   <li>Point 1</li>
   *   <li>Point 2</li>
   * </ul>
   * // Header 3 and 4 are the same with above
   * </body>
   */
  @Override
  public void writePxtSummary(Element body, XSLFTextShape placeholder) {
    log.info("Writing HTML from field [PXT Summary] to the PPT slide.");
    log.debug("HTML content to be writing:\n{}", body);
    // Clear the empty first paragraph
    placeholder.clearText();

    /* First block */
    // Write the Header
    String header1Text = body.child(0).text();
    writeHeader(header1Text, placeholder);
    // Write the List
    Element ul1 = body.child(1);
    writeBulletList(ul1, placeholder);

    /* Second block */
    // Write the Header
    String header2Text = body.child(2).text();
    writeHeader(header2Text, placeholder);
    // Write the List
    Element ul2 = body.child(3);
    writeBulletList(ul2, placeholder);

    /* Third block */
    // Write the Header
    String header3Text = body.child(4).text();
    writeHeader(header3Text, placeholder);
    // Write the List
    Element ul3 = body.child(5);
    writeBulletList(ul3, placeholder);

    /* Fourth block */
    // Write the Header
    String header4Text = body.child(6).text();
    writeHeader(header4Text, placeholder);
    // Write the List
    Element ul4 = body.child(7);
    writeBulletList(ul4, placeholder);
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

  private void writeHeader(String header, XSLFTextShape placeholder) {
    XSLFTextRun header1TextRun = placeholder.addNewTextParagraph().addNewTextRun();
    header1TextRun.setBold(true);
    header1TextRun.setText(header);
  }
}
