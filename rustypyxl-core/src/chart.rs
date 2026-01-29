//! Chart support for Excel workbooks.
//!
//! This module provides structures for creating and manipulating charts in Excel files.

use std::collections::HashMap;

/// Type of chart.
#[derive(Clone, Debug, PartialEq)]
pub enum ChartType {
    /// Bar chart (horizontal bars).
    Bar,
    /// Column chart (vertical bars).
    Column,
    /// Line chart.
    Line,
    /// Pie chart.
    Pie,
    /// Doughnut chart.
    Doughnut,
    /// Area chart.
    Area,
    /// Scatter/XY chart.
    Scatter,
    /// Radar chart.
    Radar,
    /// Stock chart.
    Stock,
    /// Surface chart.
    Surface,
    /// Bubble chart.
    Bubble,
}

impl ChartType {
    /// Get the XML type name for this chart type.
    pub fn xml_type(&self) -> &'static str {
        match self {
            ChartType::Bar => "barChart",
            ChartType::Column => "barChart",
            ChartType::Line => "lineChart",
            ChartType::Pie => "pieChart",
            ChartType::Doughnut => "doughnutChart",
            ChartType::Area => "areaChart",
            ChartType::Scatter => "scatterChart",
            ChartType::Radar => "radarChart",
            ChartType::Stock => "stockChart",
            ChartType::Surface => "surfaceChart",
            ChartType::Bubble => "bubbleChart",
        }
    }
}

/// Grouping style for bar/column charts.
#[derive(Clone, Debug, Default, PartialEq)]
pub enum BarGrouping {
    /// Clustered (side by side).
    #[default]
    Clustered,
    /// Stacked.
    Stacked,
    /// 100% stacked.
    PercentStacked,
}

/// Direction for bar charts.
#[derive(Clone, Debug, Default, PartialEq)]
pub enum BarDirection {
    /// Horizontal bars.
    Bar,
    /// Vertical bars (columns).
    #[default]
    Col,
}

/// A data series in a chart.
#[derive(Clone, Debug)]
pub struct ChartSeries {
    /// Series name/title.
    pub name: Option<String>,
    /// Reference to category (X axis) data, e.g., "Sheet1!$A$2:$A$10".
    pub categories: Option<String>,
    /// Reference to values (Y axis) data, e.g., "Sheet1!$B$2:$B$10".
    pub values: String,
    /// Fill color for the series.
    pub fill_color: Option<String>,
    /// Line color for the series.
    pub line_color: Option<String>,
    /// Marker style for line/scatter charts.
    pub marker_style: Option<String>,
}

impl ChartSeries {
    /// Create a new series with values reference.
    pub fn new<S: Into<String>>(values: S) -> Self {
        ChartSeries {
            name: None,
            categories: None,
            values: values.into(),
            fill_color: None,
            line_color: None,
            marker_style: None,
        }
    }

    /// Set the series name.
    pub fn with_name<S: Into<String>>(mut self, name: S) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Set the categories reference.
    pub fn with_categories<S: Into<String>>(mut self, categories: S) -> Self {
        self.categories = Some(categories.into());
        self
    }

    /// Set the fill color.
    pub fn with_fill_color<S: Into<String>>(mut self, color: S) -> Self {
        self.fill_color = Some(color.into());
        self
    }
}

/// Chart axis configuration.
#[derive(Clone, Debug, Default)]
pub struct ChartAxis {
    /// Axis title.
    pub title: Option<String>,
    /// Minimum value (None for auto).
    pub min_val: Option<f64>,
    /// Maximum value (None for auto).
    pub max_val: Option<f64>,
    /// Major unit/interval.
    pub major_unit: Option<f64>,
    /// Minor unit/interval.
    pub minor_unit: Option<f64>,
    /// Number format for axis labels.
    pub number_format: Option<String>,
    /// Whether to show gridlines.
    pub gridlines: bool,
    /// Position: "b" (bottom), "l" (left), "t" (top), "r" (right).
    pub position: Option<String>,
}

impl ChartAxis {
    /// Create a new axis.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the axis title.
    pub fn with_title<S: Into<String>>(mut self, title: S) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Set min and max values.
    pub fn with_range(mut self, min: f64, max: f64) -> Self {
        self.min_val = Some(min);
        self.max_val = Some(max);
        self
    }
}

/// Chart legend configuration.
#[derive(Clone, Debug, Default)]
pub struct ChartLegend {
    /// Legend position: "b" (bottom), "l" (left), "t" (top), "r" (right), "tr" (top-right).
    pub position: String,
    /// Whether to show the legend.
    pub visible: bool,
}

impl ChartLegend {
    /// Create a new legend.
    pub fn new() -> Self {
        ChartLegend {
            position: "r".to_string(),
            visible: true,
        }
    }

    /// Set the legend position.
    pub fn with_position<S: Into<String>>(mut self, position: S) -> Self {
        self.position = position.into();
        self
    }

    /// Set visibility.
    pub fn with_visible(mut self, visible: bool) -> Self {
        self.visible = visible;
        self
    }
}

/// Chart title configuration.
#[derive(Clone, Debug, Default)]
pub struct ChartTitle {
    /// Title text.
    pub text: Option<String>,
    /// Font size.
    pub font_size: Option<f64>,
    /// Font bold.
    pub bold: bool,
}

impl ChartTitle {
    /// Create a new title.
    pub fn new<S: Into<String>>(text: S) -> Self {
        ChartTitle {
            text: Some(text.into()),
            font_size: None,
            bold: false,
        }
    }
}

/// Anchor position for a chart.
#[derive(Clone, Debug)]
pub struct ChartAnchor {
    /// Cell reference for top-left corner.
    pub from_cell: String,
    /// Column offset in EMUs (English Metric Units).
    pub from_col_offset: u32,
    /// Row offset in EMUs.
    pub from_row_offset: u32,
    /// Cell reference for bottom-right corner.
    pub to_cell: Option<String>,
    /// Column offset for bottom-right.
    pub to_col_offset: u32,
    /// Row offset for bottom-right.
    pub to_row_offset: u32,
}

impl ChartAnchor {
    /// Create an anchor at the specified cell.
    pub fn at<S: Into<String>>(cell: S) -> Self {
        ChartAnchor {
            from_cell: cell.into(),
            from_col_offset: 0,
            from_row_offset: 0,
            to_cell: None,
            to_col_offset: 0,
            to_row_offset: 0,
        }
    }

    /// Set the size by specifying the to_cell.
    pub fn with_size<S: Into<String>>(mut self, to_cell: S) -> Self {
        self.to_cell = Some(to_cell.into());
        self
    }
}

/// A chart in an Excel worksheet.
#[derive(Clone, Debug)]
pub struct Chart {
    /// Chart type.
    pub chart_type: ChartType,
    /// Chart title.
    pub title: Option<ChartTitle>,
    /// Data series.
    pub series: Vec<ChartSeries>,
    /// X axis (category axis).
    pub x_axis: Option<ChartAxis>,
    /// Y axis (value axis).
    pub y_axis: Option<ChartAxis>,
    /// Legend configuration.
    pub legend: Option<ChartLegend>,
    /// Position/anchor in the worksheet.
    pub anchor: Option<ChartAnchor>,
    /// Width in EMUs (914400 EMUs = 1 inch).
    pub width: u32,
    /// Height in EMUs.
    pub height: u32,
    /// Bar grouping (for bar/column charts).
    pub bar_grouping: BarGrouping,
    /// Bar direction (for bar/column charts).
    pub bar_direction: BarDirection,
    /// Additional properties.
    pub properties: HashMap<String, String>,
}

impl Chart {
    /// Create a new chart of the specified type.
    pub fn new(chart_type: ChartType) -> Self {
        Chart {
            chart_type,
            title: None,
            series: Vec::new(),
            x_axis: None,
            y_axis: None,
            legend: Some(ChartLegend::new()),
            anchor: None,
            width: 4572000,  // ~5 inches
            height: 2743200, // ~3 inches
            bar_grouping: BarGrouping::default(),
            bar_direction: BarDirection::default(),
            properties: HashMap::new(),
        }
    }

    /// Create a bar chart.
    pub fn bar() -> Self {
        let mut chart = Self::new(ChartType::Bar);
        chart.bar_direction = BarDirection::Bar;
        chart
    }

    /// Create a column chart.
    pub fn column() -> Self {
        let mut chart = Self::new(ChartType::Column);
        chart.bar_direction = BarDirection::Col;
        chart
    }

    /// Create a line chart.
    pub fn line() -> Self {
        Self::new(ChartType::Line)
    }

    /// Create a pie chart.
    pub fn pie() -> Self {
        Self::new(ChartType::Pie)
    }

    /// Create a scatter chart.
    pub fn scatter() -> Self {
        Self::new(ChartType::Scatter)
    }

    /// Create an area chart.
    pub fn area() -> Self {
        Self::new(ChartType::Area)
    }

    /// Set the chart title.
    pub fn with_title<S: Into<String>>(mut self, title: S) -> Self {
        self.title = Some(ChartTitle::new(title));
        self
    }

    /// Add a data series.
    pub fn add_series(&mut self, series: ChartSeries) {
        self.series.push(series);
    }

    /// Set the X axis.
    pub fn with_x_axis(mut self, axis: ChartAxis) -> Self {
        self.x_axis = Some(axis);
        self
    }

    /// Set the Y axis.
    pub fn with_y_axis(mut self, axis: ChartAxis) -> Self {
        self.y_axis = Some(axis);
        self
    }

    /// Set the legend.
    pub fn with_legend(mut self, legend: ChartLegend) -> Self {
        self.legend = Some(legend);
        self
    }

    /// Set the anchor position.
    pub fn with_anchor(mut self, anchor: ChartAnchor) -> Self {
        self.anchor = Some(anchor);
        self
    }

    /// Set the chart size in inches.
    pub fn with_size_inches(mut self, width: f64, height: f64) -> Self {
        // 914400 EMUs per inch
        self.width = (width * 914400.0) as u32;
        self.height = (height * 914400.0) as u32;
        self
    }

    /// Set bar grouping style.
    pub fn with_grouping(mut self, grouping: BarGrouping) -> Self {
        self.bar_grouping = grouping;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_bar_chart() {
        let chart = Chart::bar()
            .with_title("Sales Data");

        assert_eq!(chart.chart_type, ChartType::Bar);
        assert!(chart.title.is_some());
    }

    #[test]
    fn test_create_line_chart() {
        let chart = Chart::line()
            .with_title("Trends");

        assert_eq!(chart.chart_type, ChartType::Line);
    }

    #[test]
    fn test_add_series() {
        let mut chart = Chart::bar();
        chart.add_series(
            ChartSeries::new("Sheet1!$B$2:$B$10")
                .with_name("Revenue")
                .with_categories("Sheet1!$A$2:$A$10")
        );

        assert_eq!(chart.series.len(), 1);
        assert_eq!(chart.series[0].name, Some("Revenue".to_string()));
    }

    #[test]
    fn test_chart_axes() {
        let chart = Chart::bar()
            .with_x_axis(ChartAxis::new().with_title("Categories"))
            .with_y_axis(ChartAxis::new().with_title("Values").with_range(0.0, 100.0));

        assert!(chart.x_axis.is_some());
        assert!(chart.y_axis.is_some());
        assert_eq!(chart.y_axis.as_ref().unwrap().min_val, Some(0.0));
    }

    #[test]
    fn test_chart_size() {
        let chart = Chart::pie().with_size_inches(6.0, 4.0);

        // 6 inches * 914400 EMUs/inch
        assert_eq!(chart.width, 5486400);
    }

    #[test]
    fn test_chart_type_xml() {
        assert_eq!(ChartType::Bar.xml_type(), "barChart");
        assert_eq!(ChartType::Line.xml_type(), "lineChart");
        assert_eq!(ChartType::Pie.xml_type(), "pieChart");
    }
}
