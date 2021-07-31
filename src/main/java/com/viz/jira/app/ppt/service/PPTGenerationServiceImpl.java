package com.viz.jira.app.ppt.service;

import com.atlassian.jira.config.util.JiraHome;
import com.atlassian.jira.issue.Issue;
import com.atlassian.plugin.spring.scanner.annotation.imports.ComponentImport;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import org.apache.commons.io.FileUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class PPTGenerationServiceImpl implements PPTGenerationService {

  private static final String TEMPLATE_FILE_NAME = "Template.pptx";

  private static final Logger log = LoggerFactory.getLogger(PPTGenerationServiceImpl.class);

  private final JiraHome jiraHome;

  @Autowired
  public PPTGenerationServiceImpl(@ComponentImport JiraHome jiraHome) {
    this.jiraHome = jiraHome;
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

    String fileName = issue.getKey() + ".pptx";

    String tmpFilePath = jiraHome.getHome().getAbsolutePath()
        + File.separator + "tmp"
        + File.separator + fileName;
    log.info("Creating a temporary file at: [{}]", tmpFilePath);
    File file = new File(tmpFilePath);

    FileUtils.copyInputStreamToFile(inputStream, file);
    return file;
  }
}
