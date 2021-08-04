package com.viz.jira.app.ppt.controller;

import com.atlassian.jira.issue.Issue;
import com.atlassian.jira.issue.IssueManager;
import com.atlassian.jira.permission.ProjectPermissions;
import com.atlassian.jira.security.JiraAuthenticationContext;
import com.atlassian.jira.security.PermissionManager;
import com.atlassian.jira.user.ApplicationUser;
import com.atlassian.plugin.spring.scanner.annotation.imports.ComponentImport;
import com.viz.jira.app.ppt.service.PPTGenerationService;
import java.io.File;
import java.nio.file.Files;
import javax.ws.rs.GET;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.QueryParam;
import javax.ws.rs.core.Response;
import javax.ws.rs.core.Response.ResponseBuilder;
import javax.ws.rs.core.Response.Status;
import javax.ws.rs.core.StreamingOutput;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

@Component
@Path("generate")
public class PPTGenerationController {

  private static final String PPT_MEDIA_TYPE = "application/vnd.ms-powerpoint";
  private static final String PPTX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  private static final String CONTENT_DISPOSITION_HEADER = "Content-Disposition";

  private static final Logger log = LoggerFactory.getLogger(PPTGenerationController.class);

  private final IssueManager issueManager;
  private final JiraAuthenticationContext authContext;
  private final PermissionManager permissionManager;
  private final PPTGenerationService pptGenerationService;

  @Autowired
  public PPTGenerationController(@ComponentImport IssueManager issueManager,
      @ComponentImport JiraAuthenticationContext authContext,
      @ComponentImport PermissionManager permissionManager,
      PPTGenerationService pptGenerationService) {
    this.issueManager = issueManager;
    this.authContext = authContext;
    this.permissionManager = permissionManager;
    this.pptGenerationService = pptGenerationService;
  }

  @GET
  @Produces({PPT_MEDIA_TYPE, PPTX_MEDIA_TYPE})
  public Response exportPPT(@QueryParam("issueKey") String issueKey) {
    log.info("Attempt to export PPT from issue [{}]...", issueKey);

    ApplicationUser user = authContext.getLoggedInUser();
    if (user == null) {
      log.info("User is not logged in. Response an Unauthorized status.");
      return Response.status(Status.UNAUTHORIZED).build();
    }

    Issue issue = issueManager.getIssueObject(issueKey);
    if (issue == null) {
      log.error("Error while export the PPT. The issue with key [{}] does not exist", issueKey);
      return Response.status(Status.NOT_FOUND).build();
    }

    if (hasNoViewIssuePermission(user, issue)) {
      log.warn("PPT EXPORT WARNING: User [{}] has no permission to browse issue [{}]",
          user.getUsername(), issueKey);
      return Response.status(Status.FORBIDDEN).build();
    }

    try {
      File file = pptGenerationService.generatePPT(issue);

      ResponseBuilder response = Response.ok((StreamingOutput) output -> {
        Files.copy(file.toPath(), output);
        boolean deleted = Files.deleteIfExists(file.toPath());
        if (deleted) {
          log.info("Temporary file has been deleted after download. [{}]", file.getName());
        }
      });
      String contentDispositionHeaderValue = "attachment; filename=\"" + file.getName() + "\"";
      response.header(CONTENT_DISPOSITION_HEADER, contentDispositionHeaderValue);

      return response.build();
    } catch (Exception e) {
      String message = String.format("Error while generation the PPT for issue [%s]", issueKey);
      log.error(message, e);
      return Response.serverError().build();
    }
  }

  private boolean hasNoViewIssuePermission(ApplicationUser user, Issue issue) {
    return !hasViewIssuePermission(user, issue);
  }

  private boolean hasViewIssuePermission(ApplicationUser user, Issue issue) {
    return permissionManager.hasPermission(ProjectPermissions.BROWSE_PROJECTS, issue, user);
  }
}
