package com.viz.jira.app.ppt.service;

import static com.viz.jira.app.ppt.sdo.CustomFieldName.COMMENT_BLOCK;
import static com.viz.jira.app.ppt.sdo.CustomFieldName.CONTACT;
import static com.viz.jira.app.ppt.sdo.CustomFieldName.CTA;
import static com.viz.jira.app.ppt.sdo.CustomFieldName.MILESTONES;
import static com.viz.jira.app.ppt.sdo.CustomFieldName.PXT_SUMMARY;
import static com.viz.jira.app.ppt.sdo.CustomFieldName.STATUS_FLAG2;
import static com.viz.jira.app.ppt.sdo.CustomFieldName.SW_LEAD;

import com.atlassian.jira.config.util.JiraHome;
import com.atlassian.jira.issue.CustomFieldManager;
import com.atlassian.jira.issue.Issue;
import com.atlassian.jira.issue.RendererManager;
import com.atlassian.jira.issue.customfields.impl.SelectCFType;
import com.atlassian.jira.issue.customfields.impl.UserCFType;
import com.atlassian.jira.issue.customfields.option.Option;
import com.atlassian.jira.issue.fields.CustomField;
import com.atlassian.jira.issue.status.Status;
import com.atlassian.jira.user.ApplicationUser;
import com.atlassian.plugin.spring.scanner.annotation.imports.ComponentImport;
import com.viz.jira.app.ppt.sdo.OverallHealthColor;
import com.viz.jira.app.ppt.sdo.ShapeName;
import com.viz.jira.app.ppt.sdo.StatusColor;
import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import javax.annotation.Nullable;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class PPTGenerationServiceImpl implements PPTGenerationService {

  private static final String TEMPLATE_FILE_NAME = "Template.pptx";

  private static final Logger log = LoggerFactory.getLogger(PPTGenerationServiceImpl.class);

  private final JiraHome jiraHome;
  private final CustomFieldManager customFieldManager;
  private final RendererManager rendererManager;
  private final HtmlToPptService htmlToPptService;

  @Autowired
  public PPTGenerationServiceImpl(@ComponentImport JiraHome jiraHome,
      @ComponentImport CustomFieldManager customFieldManager,
      @ComponentImport RendererManager rendererManager,
      HtmlToPptService htmlToPptService) {
    this.jiraHome = jiraHome;
    this.customFieldManager = customFieldManager;
    this.rendererManager = rendererManager;
    this.htmlToPptService = htmlToPptService;
  }

  @Override
  public File generatePPT(Issue issue) throws IOException {
    log.info("Generating PPT for issue [{}]...", issue.getKey());
    InputStream inputStream = Thread.currentThread().getContextClassLoader()
        .getResourceAsStream(TEMPLATE_FILE_NAME);

    if (inputStream == null) {
      String message = String.format(
          "Error while generating the PPT for issue [%s]. The PPT template file could not be read.",
          issue.getKey());
      throw new IOException(message);
    }

    XMLSlideShow ppt = new XMLSlideShow(inputStream);

    /* The Template contains one slide already (The PPT is created by me, so I know it) */
    XSLFSlide slide = ppt.getSlides().get(0);

    List<XSLFShape> shapes = slide.getShapes();
    for (XSLFShape shape : shapes) {
      writeIssueDataToShape(issue, shape);
    }

    String tmpFilePath = getTempFilePath(issue);

    // Write the new PPT to a new temporary file
    log.info("Writing PPT data to a temporary file at: [{}]", tmpFilePath);
    FileOutputStream out = new FileOutputStream(tmpFilePath);
    ppt.write(out);
    out.close();
    ppt.close();
    log.info("The PPT data has been successfully written to [{}]", tmpFilePath);

    // Get and return the tmp file after writing
    return new File(tmpFilePath);
  }

  private void writeIssueDataToShape(Issue issue, XSLFShape shape) {
    String shapeName = shape.getShapeName();
    // We have 3 XSLFTable in the Template (Top Table, Left Table and Right Table)
    if (shape instanceof XSLFTable) {
      writeIssueDataToTheTable(issue, (XSLFTable) shape);
    } else {
      log.warn("Found an unknown shape [{}] in the template. Shape type: [{}]",
          shapeName, shape.getClass().getName());
    }
  }

  private String getTempFilePath(Issue issue) {
    String fileName = getExportFileName(issue);

    return jiraHome.getHome().getAbsolutePath()
        + File.separator + "tmp" + File.separator + fileName;
  }

  private String getExportFileName(Issue issue) {
    return issue.getKey() + ".pptx";
  }

  private void writeIssueDataToTheTable(Issue issue, XSLFTable table) {
    String tableName = table.getShapeName();
    log.info("Start writing Issue data to the table [{}]", tableName);

    switch (tableName) {
      case ShapeName.TOP_TABLE:
        writeTopTable(issue, table);
        break;
      case ShapeName.LEFT_TABLE:
        writeLeftTable(issue, table);
        break;
      case ShapeName.RIGHT_TABLE:
        writeRightTable(issue, table);
        break;
      default:
        log.warn("The table with name [{}] is not yet handled. ", tableName);
    }
  }

  private void writeRightTable(Issue issue, XSLFTable table) {
    /* "Milestones" in Jira to "Timeline/Milestones" in slide */
    CustomField milestonesField = getFirstCustomFieldByName(MILESTONES);
    if (milestonesField != null) {
      String htmlValue = exportHtmlValueFromMultiLineTextField(milestonesField, issue);
      Document document = Jsoup.parse(htmlValue);
      log.info("{} parsed HTML value:\n{}", MILESTONES, document);
      Elements tables = document.body().getElementsByTag("table");

      if (!tables.isEmpty()) {
        Element htmlTable = tables.first();
        writeMilestonesTable(issue, htmlTable, table);
      } else {
        log.warn("There is no milestones table found.");
      }
    }
  }

  private void writeMilestonesTable(Issue issue, Element htmlTable, XSLFTable table) {
    /* The XSLFTable template has only 17 rows to write data (plus 1 for header). If the customer
     * input more than 17 milestones, the rest will be ignored. In that case, the developers need
     * to edit the PPT template. */
    Element tableBody = htmlTable.getElementsByTag("tbody").get(0);
    Elements htmlRows = tableBody.children();

    // Log a warning if the Milestones table has more than 17 milestones (plus 1 header)
    if (htmlRows.size() > 18) {
      log.warn("The Milestone table on issue [{}] has more than 17 item, and it exceed the number "
              + "of item the PPT template can hold. The exceeded items will be ignored.",
          issue.getKey());
    }

    if (htmlRows.isEmpty()) {
      log.warn("The table Milestones on issue [{}] has no row.", issue.getKey());
      return;
    }

    /* Start writing from the second row, we don't need the header */
    for (int rowIndex = 1; rowIndex < htmlRows.size(); rowIndex++) {
      /* We only write the first 17 rows to the PPT, since the Template has only 17 placeholders */
      if (rowIndex > 17) {
        break;
      }

      Element htmlRow = htmlRows.get(rowIndex);

      // The 1st column from Jira -> The 1st column in PPT
      String planned = htmlRow.child(0).text();
      XSLFTableCell pptCell = table.getCell(rowIndex, 0);
      setTextKeepFormat(planned, pptCell);

      // The 3rd column from Jira -> The 2nd column in PPT
      String status = htmlRow.child(2).text();
      pptCell = table.getCell(rowIndex, 1);
      setTextKeepFormat(status, pptCell);

      // The 5th column from Jira -> The 3rd column in PPT
      String task = htmlRow.child(4).text();
      pptCell = table.getCell(rowIndex, 2);
      setTextKeepFormat(task, pptCell);
    }
  }

  private void writeLeftTable(Issue issue, XSLFTable table) {
    writePxtSummaryToTable(issue, table);
    writeExternalOwner(issue, table);
    writeInternalOwner(issue, table);
    writeStatusUpdate(issue, table);
  }

  private void writeStatusUpdate(Issue issue, XSLFTable table) {
    /* "Comment Block" in Jira to "Status Update and Issues/Risks" in slide */
    CustomField commentBlockField = getFirstCustomFieldByName(COMMENT_BLOCK);
    XSLFTableCell commentBlockCell = table.getCell(11, 0);
    if (commentBlockField != null) {
      String htmlValue = exportHtmlValueFromMultiLineTextField(commentBlockField, issue);
      Document document = Jsoup.parse(htmlValue);
      log.info("{} parsed value:\n{}", COMMENT_BLOCK, document);
      Element body = document.body();
      htmlToPptService.writeHtmlToTextShape(body, commentBlockCell);
    } else {
      commentBlockCell.clearText();
    }
  }

  private void writeInternalOwner(Issue issue, XSLFTable table) {
    /* The "CTA" and "SW Lead" in Jira will be written to "Internal Owner" in the slide. */
    List<String> names = new ArrayList<>();

    CustomField ctaField = getFirstCustomFieldByName(CTA);
    if (ctaField != null) {
      String ctaUserName = getUserDisplayNameUserPickerField(ctaField, issue);
      if (ctaUserName != null) {
        names.add(ctaUserName);
      }
    }

    CustomField swLeadField = getFirstCustomFieldByName(SW_LEAD);
    if (swLeadField != null) {
      String swLeadUserName = getUserDisplayNameUserPickerField(swLeadField, issue);
      if (swLeadUserName != null) {
        names.add(swLeadUserName);
      }
    }

    String internalOwnerNames = String.join(", ", names);
    XSLFTableCell internalOwnerCell = table.getCell(9, 1);
    setTextKeepFormat(internalOwnerNames, internalOwnerCell);
  }

  private void writeExternalOwner(Issue issue, XSLFTable table) {
    /* The "Contact" value in Jira will be written to "External Owner" in the slide. */
    XSLFTableCell externalOwnerCell = table.getCell(9, 0);
    CustomField contactField = getFirstCustomFieldByName(CONTACT);
    if (contactField != null) {
      String userName = getUserDisplayNameUserPickerField(contactField, issue);
      log.info("[{}] value from issue [{}]: [{}]", CONTACT, issue.getKey(), userName);
      setTextKeepFormat(userName, externalOwnerCell);
    } else {
      externalOwnerCell.clearText();
    }
  }

  private void writePxtSummaryToTable(Issue issue, XSLFTable table) {
    log.info("Writing [{}] to PPT", PXT_SUMMARY);
    CustomField pxtSummaryField = getFirstCustomFieldByName(PXT_SUMMARY);
    if (pxtSummaryField != null) {
      String htmlValue = exportHtmlValueFromMultiLineTextField(pxtSummaryField, issue);
      Document document = Jsoup.parse(htmlValue);
      log.info("{} parsed value:\n{}", PXT_SUMMARY, document);
      Element pxtBody = document.body();

      // Write description
      XSLFTableCell descriptionCell = table.getCell(1, 0);
      Element description = pxtBody.child(1);
      htmlToPptService.writeHtmlToTextShape(description, descriptionCell);

      // Write product intercepts
      XSLFTableCell productInterceptsCell = table.getCell(3, 0);
      Element productIntercepts = pxtBody.child(3);
      htmlToPptService.writeHtmlToTextShape(productIntercepts, productInterceptsCell);

      // Write Success Metric
      XSLFTableCell successMetricCell = table.getCell(5, 0);
      Element successMetric = pxtBody.child(5);
      htmlToPptService.writeHtmlToTextShape(successMetric, successMetricCell);

      // Write Key Deliverables
      XSLFTableCell keyDeliverablesCell = table.getCell(7, 0);
      Element keyDeliverables = pxtBody.child(7);
      htmlToPptService.writeHtmlToTextShape(keyDeliverables, keyDeliverablesCell);
    }
  }

  private String getUserDisplayNameUserPickerField(CustomField customField, Issue issue) {
    UserCFType cfType = (UserCFType) customField.getCustomFieldType();
    ApplicationUser user = cfType.getValueFromIssue(customField, issue);
    return user != null ? user.getDisplayName() : null;
  }

  private String exportHtmlValueFromMultiLineTextField(CustomField customField, Issue issue) {
    String wikiMarkupValue = getTextCustomFieldValue(customField, issue);
    return wikiMarkupToHtml(wikiMarkupValue, issue);
  }

  private String wikiMarkupToHtml(String markup, Issue issue) {
    String rendererType = "atlassian-wiki-renderer";
    return rendererManager.getRenderedContent(rendererType, markup, issue.getIssueRenderContext());
  }

  private String getTextCustomFieldValue(CustomField customField, Issue issue) {
    String value = customField.getValueFromIssue(issue);
    return value != null ? value : "";
  }

  private String getOverallHealth(Issue issue) {
    /* The "Status-Flag2" value in Jira will be written to "Overall Health" in the slide. */
    CustomField statusFlag2Field = getFirstCustomFieldByName(STATUS_FLAG2);
    if (statusFlag2Field != null) {
      return getSingleSelectValue(statusFlag2Field, issue);
    }
    return "";
  }

  /**
   * Get the text value from a single select custom field. If the value is not set, we will return
   * an empty string "".
   *
   * @param customField The custom field we want to get value from
   * @param issue The issue we want to get value from
   * @return Custom field value, or "" if there is no value
   */
  private String getSingleSelectValue(CustomField customField, Issue issue) {
    SelectCFType cfType = (SelectCFType) customField.getCustomFieldType();
    Option option = cfType.getValueFromIssue(customField, issue);
    return option != null ? option.getValue() : "";
  }

  private void writeTopTable(Issue issue, XSLFTable table) {
    // Edit issue key
    log.info("Writing Issue Key: [{}]", issue.getKey());
    XSLFTableCell issueKeyCell = table.getCell(0, 0);
    setTextKeepFormat(issue.getKey(), issueKeyCell);

    // Edit issue summary
    log.info("Writing Issue Summary: [{}]", issue.getSummary());
    XSLFTableCell summaryCell = table.getCell(1, 0);
    setTextKeepFormat(issue.getSummary(), summaryCell);

    // Edit Date (Updated)
    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
    String updated = issue.getUpdated().toLocalDateTime().format(formatter);
    log.info("Writing Issue updated date: [{}]", updated);
    XSLFTableCell dateCell = table.getCell(1, 1);
    setTextKeepFormat(updated, dateCell);

    // Edit Phase (Status)
    writeEditPhase(issue, table);

    // Edit Overall Health (Status-Flag2)
    writeOverallHealth(issue, table);
  }

  private void writeOverallHealth(Issue issue, XSLFTable table) {
    String overallHealth = getOverallHealth(issue);
    log.info("Writing Overall Health: [{}]", overallHealth);
    XSLFTableCell overallHealthCell = table.getCell(1, 3);
    setTextKeepFormat(overallHealth, overallHealthCell);

    Color color = OverallHealthColor.get(overallHealth);
    if (color != null) {
      overallHealthCell.setFillColor(color);
    } else {
      log.error("No color matching for Status-Flag2 value [{}]", overallHealth);
    }
  }

  private void writeEditPhase(Issue issue, XSLFTable table) {
    Status status = issue.getStatus();
    String statusName = status.getName().toUpperCase();
    log.info("Writing Issue status: [{}]", statusName);
    XSLFTableCell phaseCell = table.getCell(1, 2);
    setTextKeepFormat(statusName, phaseCell);

    // Set color for the cell based on status color
    String colorKey = status.getStatusCategory().getKey();
    Color color = StatusColor.getColor(colorKey);
    phaseCell.setFillColor(color);
  }

  private void setTextKeepFormat(String text, XSLFTableCell tableCell) {
    if (tableCell != null) {
      // Assume the cell has only one paragraph
      XSLFTextParagraph paragraph = tableCell.getTextParagraphs().get(0);
      List<XSLFTextRun> textRuns = paragraph.getTextRuns();

      /* Making sure that the paragraph has only one text run, we will remove all other text runs
       * after index 0 (from the first text run) */
      if (textRuns.size() > 1) {
        textRuns.subList(1, textRuns.size()).clear();
      }

      /* Get the remaining text run (at position 0) and set text for it. */
      XSLFTextRun textRun = textRuns.get(0);
      textRun.setText(text);
    } else {
      log.error("The Table Cell is null. Could not write data");
    }
  }

  @Nullable
  private CustomField getFirstCustomFieldByName(String fieldName) {
    Collection<CustomField> fields = customFieldManager.getCustomFieldObjectsByName(fieldName);

    if (fields.isEmpty()) {
      log.warn("There is no field with name [{}] in the instance.", fieldName);
      return null;
    }

    // Inform the users about field with the same name
    if (fields.size() > 1) {
      log.warn("Found {} fields with the same name [{}] in the instance. "
          + "Will select the first field.", fields.size(), fieldName);
    }

    return (CustomField) fields.toArray()[0];
  }
}
