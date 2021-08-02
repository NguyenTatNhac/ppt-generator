package com.viz.jira.app.ppt.service;

import static com.viz.jira.app.ppt.sdo.CustomFieldName.CONTACT;
import static com.viz.jira.app.ppt.sdo.CustomFieldName.CTA;
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
import com.atlassian.jira.user.ApplicationUser;
import com.atlassian.plugin.spring.scanner.annotation.imports.ComponentImport;
import com.viz.jira.app.ppt.sdo.SlidePlaceholderName;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import javax.annotation.Nullable;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
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

    XSLFTextShape[] placeholders = slide.getPlaceholders();
    for (XSLFTextShape placeholder : placeholders) {
      writeIssueDataToThePlaceholder(issue, placeholder);
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

  private String getTempFilePath(Issue issue) {
    String fileName = getExportFileName(issue);

    return jiraHome.getHome().getAbsolutePath()
        + File.separator + "tmp" + File.separator + fileName;
  }

  private String getExportFileName(Issue issue) {
    return issue.getKey() + ".pptx";
  }

  private void writeIssueDataToThePlaceholder(Issue issue, XSLFTextShape placeholder) {
    String placeholderName = placeholder.getShapeName();
    switch (placeholderName) {
      case SlidePlaceholderName.KEY:
        writeIssueKeyToThePlaceholder(issue, placeholder);
        break;
      case SlidePlaceholderName.SUMMARY:
        writeSummaryToThePlaceholder(issue, placeholder);
        break;
      case SlidePlaceholderName.DATE:
        writeDateToThePlaceholder(issue, placeholder);
        break;
      case SlidePlaceholderName.PHASE:
        writePhaseToThePlaceholder(issue, placeholder);
        break;
      case SlidePlaceholderName.OVERALL_HEALTH:
        writeOverallHealthToThePlaceholder(issue, placeholder);
        break;
      case SlidePlaceholderName.PXT_SUMMARY:
        writePxtSummaryToThePlaceholder(issue, placeholder);
        break;
      case SlidePlaceholderName.EXTERNAL_OWNER:
        writeExternalOwnerToThePlaceholder(issue, placeholder);
        break;
      case SlidePlaceholderName.INTERNAL_OWNER:
        writeInternalOwnerToThePlaceholder(issue, placeholder);
        break;
      default:
        log.warn("The placeholder [{}] is not yet handled. "
            + "Please contact [nguyentatnhac@gmail.com] to get support.", placeholderName);
    }
  }

  private void writeInternalOwnerToThePlaceholder(Issue issue, XSLFTextShape placeholder) {
    /* According to the requirement: The "CTA" and "SW Lead" values in Jira will be written to
     * "Internal Owner" in the slide. */
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
    writeToTextPlaceholder(internalOwnerNames, placeholder);
  }

  private void writeExternalOwnerToThePlaceholder(Issue issue, XSLFTextShape placeholder) {
    /* According to the requirement: The "Contact" value in Jira will be written to "External Owner"
     * in the slide. */
    CustomField contactField = getFirstCustomFieldByName(CONTACT);

    // Inform the users about field with the same name
    if (contactField != null) {
      String userName = getUserDisplayNameUserPickerField(contactField, issue);
      writeToTextPlaceholder(userName, placeholder);
    }
  }

  private String getUserDisplayNameUserPickerField(CustomField customField, Issue issue) {
    UserCFType cfType = (UserCFType) customField.getCustomFieldType();
    ApplicationUser user = cfType.getValueFromIssue(customField, issue);
    return user != null ? user.getDisplayName() : null;
  }

  private void writePxtSummaryToThePlaceholder(Issue issue, XSLFTextShape placeholder) {
    CustomField pxtSummaryField = getFirstCustomFieldByName(PXT_SUMMARY);
    if (pxtSummaryField != null) {
      String htmlValue = exportHtmlValueFromMultiLineTextField(pxtSummaryField, issue);
      Document document = Jsoup.parse(htmlValue);
      htmlToPptService.writePxtSummary(document, placeholder);
    }
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

  private void writeOverallHealthToThePlaceholder(Issue issue, XSLFTextShape placeholder) {
    /* According to the requirement: The "Status-Flag2" value in Jira will be written to "Overall
     * Health" in the slide. */
    CustomField statusFlag2Field = getFirstCustomFieldByName(STATUS_FLAG2);
    if (statusFlag2Field != null) {
      String statusFlag2Value = getSingleSelectValue(statusFlag2Field, issue);
      writeToTextPlaceholder(statusFlag2Value, placeholder);
    }
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

  private void writePhaseToThePlaceholder(Issue issue, XSLFTextShape placeholder) {
    /* According to the requirement: The issue status will be written to "Phase" in the slide. */
    String status = issue.getStatus().getName();
    writeToTextPlaceholder(status, placeholder);
  }

  private void writeDateToThePlaceholder(Issue issue, XSLFTextShape placeholder) {
    /* According to the requirement: The issue updated date will be written to "Date" in the slide. */
    String updatedTime = issue.getUpdated().toString();
    writeToTextPlaceholder(updatedTime, placeholder);
  }

  private void writeSummaryToThePlaceholder(Issue issue, XSLFTextShape placeholder) {
    writeToTextPlaceholder(issue.getSummary(), placeholder);
  }

  private void writeIssueKeyToThePlaceholder(Issue issue, XSLFTextShape placeholder) {
    writeToTextPlaceholder(issue.getKey(), placeholder);
  }

  private void writeToTextPlaceholder(String text, XSLFTextShape placeholder) {
    String textToSet = text != null ? text : "";
    placeholder.setText(textToSet);
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
