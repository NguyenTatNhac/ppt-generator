package com.viz.jira.app.ppt.sdo;

import com.atlassian.jira.issue.status.category.StatusCategory;
import java.awt.Color;

public class StatusColor {

  public static final Color TO_DO = new Color(66, 82, 110);
  public static final Color IN_PROGRESS = new Color(0, 82, 204);
  public static final Color COMPLETE = new Color(0, 135, 90);

  public static Color getColor(String colorKey) {
    switch (colorKey) {
      case StatusCategory.TO_DO:
        return TO_DO;
      case StatusCategory.IN_PROGRESS:
        return IN_PROGRESS;
      case StatusCategory.COMPLETE:
        return COMPLETE;
      default:
        return Color.gray;
    }
  }

  private StatusColor() {
    // Util
  }
}
