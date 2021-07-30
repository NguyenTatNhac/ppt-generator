package com.viz.jira.app.ppt.service;

import com.atlassian.jira.issue.Issue;
import java.io.File;
import java.io.IOException;

public interface PPTGenerationService {

  File generatePPT(Issue issue) throws IOException;
}
