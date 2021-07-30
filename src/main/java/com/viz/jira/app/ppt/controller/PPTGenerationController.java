package com.viz.jira.app.ppt.controller;

import com.atlassian.jira.issue.Issue;
import com.atlassian.jira.issue.IssueManager;
import com.atlassian.jira.security.JiraAuthenticationContext;
import com.atlassian.jira.user.ApplicationUser;
import com.atlassian.plugin.spring.scanner.annotation.imports.ComponentImport;
import com.viz.jira.app.ppt.service.PPTGenerationService;
import java.io.File;
import java.io.IOException;
import javax.ws.rs.GET;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.QueryParam;
import javax.ws.rs.core.Response;
import javax.ws.rs.core.Response.ResponseBuilder;
import javax.ws.rs.core.Response.Status;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

@Component
@Path("generate")
public class PPTGenerationController {

  private static final String PPT_MEDIA_TYPE = "application/vnd.ms-powerpoint";
  private static final String PPTX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

  private static final Logger log = LoggerFactory.getLogger(PPTGenerationController.class);

  private final IssueManager issueManager;
  private final JiraAuthenticationContext authContext;
  private final PPTGenerationService pptGenerationService;

  @Autowired
  public PPTGenerationController(@ComponentImport IssueManager issueManager,
      @ComponentImport JiraAuthenticationContext authContext,
      PPTGenerationService pptGenerationService) {
    this.issueManager = issueManager;
    this.authContext = authContext;
    this.pptGenerationService = pptGenerationService;
  }

  @GET
  @Produces({PPT_MEDIA_TYPE, PPTX_MEDIA_TYPE})
  public Response exportPPT(@QueryParam("issueKey") String issueKey) {
    log.warn("Exporting PPT from issue [{}]...", issueKey);
    ApplicationUser user = authContext.getLoggedInUser();

    if (user == null) {
      log.warn("User is not logged in");
      return Response.status(Status.UNAUTHORIZED).build();
    }

    Issue issue = issueManager.getIssueObject(issueKey);

    if (issue == null) {
      return Response.status(Status.NOT_FOUND).build();
    }

    try {
      File file = pptGenerationService.generatePPT(issue);

      ResponseBuilder response = Response.ok(file);
      response.header("Content-Disposition", "attachment; filename=\"" + file.getName() + "\"");

      return response.build();
    } catch (IOException e) {
      e.printStackTrace();
      return Response.serverError().build();
    }
  }
}
