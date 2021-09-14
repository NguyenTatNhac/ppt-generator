package com.viz.jira.app.ppt.sdo;

import java.awt.Color;

public class OverallHealthColor {

  public static final Color GREEN = new Color(0, 135, 90);
  public static final Color YELLOW = new Color(255, 192, 0);

  public static Color get(String color) {
    switch (color.toLowerCase()) {
      case "red":
        return Color.RED;
      case "yellow":
        return YELLOW;
      case "green":
        return GREEN;
      default:
        return null;
    }
  }

  private OverallHealthColor() {
    // Util
  }
}
