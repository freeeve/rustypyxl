//! Conditional formatting support for Excel workbooks.
//!
//! This module provides structures for creating and managing conditional formatting rules.

/// Type of conditional formatting rule.
#[derive(Clone, Debug, PartialEq)]
pub enum ConditionalFormatType {
    /// Format based on cell value.
    CellIs,
    /// Format based on formula result.
    Expression,
    /// Color scale (gradient).
    ColorScale,
    /// Data bar.
    DataBar,
    /// Icon set.
    IconSet,
    /// Top/bottom N values.
    Top10,
    /// Above/below average.
    AboveAverage,
    /// Duplicate values.
    DuplicateValues,
    /// Unique values.
    UniqueValues,
    /// Text contains.
    ContainsText,
    /// Text does not contain.
    NotContainsText,
    /// Text begins with.
    BeginsWith,
    /// Text ends with.
    EndsWith,
    /// Contains blanks.
    ContainsBlanks,
    /// Contains no blanks.
    NotContainsBlanks,
    /// Contains errors.
    ContainsErrors,
    /// Contains no errors.
    NotContainsErrors,
    /// Time period (today, yesterday, this week, etc.).
    TimePeriod,
}

/// Operator for cell value comparisons.
#[derive(Clone, Debug, PartialEq)]
pub enum ConditionalOperator {
    /// Less than.
    LessThan,
    /// Less than or equal.
    LessThanOrEqual,
    /// Equal.
    Equal,
    /// Not equal.
    NotEqual,
    /// Greater than or equal.
    GreaterThanOrEqual,
    /// Greater than.
    GreaterThan,
    /// Between (inclusive).
    Between,
    /// Not between.
    NotBetween,
}

impl ConditionalOperator {
    /// Get the XML attribute value for this operator.
    pub fn xml_value(&self) -> &'static str {
        match self {
            ConditionalOperator::LessThan => "lessThan",
            ConditionalOperator::LessThanOrEqual => "lessThanOrEqual",
            ConditionalOperator::Equal => "equal",
            ConditionalOperator::NotEqual => "notEqual",
            ConditionalOperator::GreaterThanOrEqual => "greaterThanOrEqual",
            ConditionalOperator::GreaterThan => "greaterThan",
            ConditionalOperator::Between => "between",
            ConditionalOperator::NotBetween => "notBetween",
        }
    }
}

/// Color for conditional formatting.
#[derive(Clone, Debug)]
pub struct ConditionalColor {
    /// RGB color value (e.g., "FF0000" for red).
    pub rgb: Option<String>,
    /// Theme color index.
    pub theme: Option<u32>,
    /// Tint value (-1.0 to 1.0).
    pub tint: Option<f64>,
}

impl ConditionalColor {
    /// Create a color from RGB hex string.
    pub fn rgb<S: Into<String>>(color: S) -> Self {
        ConditionalColor {
            rgb: Some(color.into()),
            theme: None,
            tint: None,
        }
    }

    /// Create a color from theme index.
    pub fn theme(index: u32) -> Self {
        ConditionalColor {
            rgb: None,
            theme: Some(index),
            tint: None,
        }
    }

    /// Common colors.
    pub fn red() -> Self {
        Self::rgb("FFFF0000")
    }

    pub fn green() -> Self {
        Self::rgb("FF00FF00")
    }

    pub fn yellow() -> Self {
        Self::rgb("FFFFFF00")
    }

    pub fn blue() -> Self {
        Self::rgb("FF0000FF")
    }

    pub fn white() -> Self {
        Self::rgb("FFFFFFFF")
    }
}

/// Color scale configuration (2 or 3 colors).
#[derive(Clone, Debug)]
pub struct ColorScale {
    /// Minimum value color.
    pub min_color: ConditionalColor,
    /// Middle value color (for 3-color scale).
    pub mid_color: Option<ConditionalColor>,
    /// Maximum value color.
    pub max_color: ConditionalColor,
    /// Minimum value type: "min", "num", "percent", "formula", "percentile".
    pub min_type: String,
    /// Minimum value (if type is "num", "percent", etc.).
    pub min_value: Option<String>,
    /// Middle value type.
    pub mid_type: Option<String>,
    /// Middle value.
    pub mid_value: Option<String>,
    /// Maximum value type.
    pub max_type: String,
    /// Maximum value.
    pub max_value: Option<String>,
}

impl ColorScale {
    /// Create a 2-color scale (min to max).
    pub fn two_color(min_color: ConditionalColor, max_color: ConditionalColor) -> Self {
        ColorScale {
            min_color,
            mid_color: None,
            max_color,
            min_type: "min".to_string(),
            min_value: None,
            mid_type: None,
            mid_value: None,
            max_type: "max".to_string(),
            max_value: None,
        }
    }

    /// Create a 3-color scale (min, mid, max).
    pub fn three_color(
        min_color: ConditionalColor,
        mid_color: ConditionalColor,
        max_color: ConditionalColor,
    ) -> Self {
        ColorScale {
            min_color,
            mid_color: Some(mid_color),
            max_color,
            min_type: "min".to_string(),
            min_value: None,
            mid_type: Some("percentile".to_string()),
            mid_value: Some("50".to_string()),
            max_type: "max".to_string(),
            max_value: None,
        }
    }

    /// Red-Yellow-Green scale.
    pub fn red_yellow_green() -> Self {
        Self::three_color(
            ConditionalColor::rgb("FFF8696B"),
            ConditionalColor::rgb("FFFFEB84"),
            ConditionalColor::rgb("FF63BE7B"),
        )
    }

    /// Green-Yellow-Red scale (reverse).
    pub fn green_yellow_red() -> Self {
        Self::three_color(
            ConditionalColor::rgb("FF63BE7B"),
            ConditionalColor::rgb("FFFFEB84"),
            ConditionalColor::rgb("FFF8696B"),
        )
    }
}

/// Data bar configuration.
#[derive(Clone, Debug)]
pub struct DataBar {
    /// Bar fill color.
    pub fill_color: ConditionalColor,
    /// Bar border color.
    pub border_color: Option<ConditionalColor>,
    /// Minimum value type.
    pub min_type: String,
    /// Minimum value.
    pub min_value: Option<String>,
    /// Maximum value type.
    pub max_type: String,
    /// Maximum value.
    pub max_value: Option<String>,
    /// Show value in cell.
    pub show_value: bool,
    /// Gradient fill.
    pub gradient: bool,
    /// Negative bar color.
    pub negative_color: Option<ConditionalColor>,
}

impl DataBar {
    /// Create a new data bar with default settings.
    pub fn new() -> Self {
        DataBar {
            fill_color: ConditionalColor::rgb("FF638EC6"),
            border_color: None,
            min_type: "min".to_string(),
            min_value: None,
            max_type: "max".to_string(),
            max_value: None,
            show_value: true,
            gradient: true,
            negative_color: None,
        }
    }

    /// Set the fill color.
    pub fn with_color(mut self, color: ConditionalColor) -> Self {
        self.fill_color = color;
        self
    }

    /// Hide the cell value.
    pub fn hide_value(mut self) -> Self {
        self.show_value = false;
        self
    }

    /// Use solid fill instead of gradient.
    pub fn solid_fill(mut self) -> Self {
        self.gradient = false;
        self
    }
}

impl Default for DataBar {
    fn default() -> Self {
        Self::new()
    }
}

/// Icon set style.
#[derive(Clone, Debug, PartialEq)]
pub enum IconSetStyle {
    /// 3 arrows (up, right, down).
    ThreeArrows,
    /// 3 arrows gray.
    ThreeArrowsGray,
    /// 3 flags.
    ThreeFlags,
    /// 3 traffic lights.
    ThreeTrafficLights,
    /// 3 signs.
    ThreeSigns,
    /// 3 symbols.
    ThreeSymbols,
    /// 3 symbols 2.
    ThreeSymbols2,
    /// 4 arrows.
    FourArrows,
    /// 4 arrows gray.
    FourArrowsGray,
    /// 4 rating.
    FourRating,
    /// 4 traffic lights.
    FourTrafficLights,
    /// 5 arrows.
    FiveArrows,
    /// 5 arrows gray.
    FiveArrowsGray,
    /// 5 rating.
    FiveRating,
    /// 5 quarters.
    FiveQuarters,
}

impl IconSetStyle {
    /// Get the XML type name.
    pub fn xml_type(&self) -> &'static str {
        match self {
            IconSetStyle::ThreeArrows => "3Arrows",
            IconSetStyle::ThreeArrowsGray => "3ArrowsGray",
            IconSetStyle::ThreeFlags => "3Flags",
            IconSetStyle::ThreeTrafficLights => "3TrafficLights1",
            IconSetStyle::ThreeSigns => "3Signs",
            IconSetStyle::ThreeSymbols => "3Symbols",
            IconSetStyle::ThreeSymbols2 => "3Symbols2",
            IconSetStyle::FourArrows => "4Arrows",
            IconSetStyle::FourArrowsGray => "4ArrowsGray",
            IconSetStyle::FourRating => "4Rating",
            IconSetStyle::FourTrafficLights => "4TrafficLights",
            IconSetStyle::FiveArrows => "5Arrows",
            IconSetStyle::FiveArrowsGray => "5ArrowsGray",
            IconSetStyle::FiveRating => "5Rating",
            IconSetStyle::FiveQuarters => "5Quarters",
        }
    }
}

/// Icon set configuration.
#[derive(Clone, Debug)]
pub struct IconSet {
    /// Icon set style.
    pub style: IconSetStyle,
    /// Show icon only (hide value).
    pub show_value: bool,
    /// Reverse icon order.
    pub reverse: bool,
    /// Custom thresholds (percentages by default).
    pub thresholds: Vec<(String, String)>, // (type, value)
}

impl IconSet {
    /// Create a new icon set.
    pub fn new(style: IconSetStyle) -> Self {
        IconSet {
            style,
            show_value: true,
            reverse: false,
            thresholds: Vec::new(),
        }
    }

    /// Hide the cell value.
    pub fn icon_only(mut self) -> Self {
        self.show_value = false;
        self
    }

    /// Reverse the icon order.
    pub fn reversed(mut self) -> Self {
        self.reverse = true;
        self
    }
}

/// Format to apply when condition is met.
#[derive(Clone, Debug, Default)]
pub struct ConditionalFormat {
    /// Font color.
    pub font_color: Option<ConditionalColor>,
    /// Bold font.
    pub bold: Option<bool>,
    /// Italic font.
    pub italic: Option<bool>,
    /// Underline.
    pub underline: Option<bool>,
    /// Strikethrough.
    pub strikethrough: Option<bool>,
    /// Fill/background color.
    pub fill_color: Option<ConditionalColor>,
    /// Border color.
    pub border_color: Option<ConditionalColor>,
    /// Number format.
    pub number_format: Option<String>,
}

impl ConditionalFormat {
    /// Create a new format.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set font color.
    pub fn with_font_color(mut self, color: ConditionalColor) -> Self {
        self.font_color = Some(color);
        self
    }

    /// Set bold.
    pub fn with_bold(mut self, bold: bool) -> Self {
        self.bold = Some(bold);
        self
    }

    /// Set fill color.
    pub fn with_fill(mut self, color: ConditionalColor) -> Self {
        self.fill_color = Some(color);
        self
    }
}

/// A conditional formatting rule.
#[derive(Clone, Debug)]
pub struct ConditionalRule {
    /// Rule type.
    pub rule_type: ConditionalFormatType,
    /// Priority (lower = higher priority).
    pub priority: u32,
    /// Operator for cell value rules.
    pub operator: Option<ConditionalOperator>,
    /// First formula/value.
    pub formula1: Option<String>,
    /// Second formula/value (for between operators).
    pub formula2: Option<String>,
    /// Text to match (for text rules).
    pub text: Option<String>,
    /// Format to apply.
    pub format: Option<ConditionalFormat>,
    /// Color scale configuration.
    pub color_scale: Option<ColorScale>,
    /// Data bar configuration.
    pub data_bar: Option<DataBar>,
    /// Icon set configuration.
    pub icon_set: Option<IconSet>,
    /// Stop if true.
    pub stop_if_true: bool,
    /// Bottom N instead of top N.
    pub bottom: bool,
    /// Percent instead of count.
    pub percent: bool,
    /// Rank/count for top/bottom rules.
    pub rank: Option<u32>,
    /// Above average (true) or below (false).
    pub above_average: bool,
    /// Equal to average allowed.
    pub equal_average: bool,
    /// Standard deviation value.
    pub std_dev: Option<u32>,
    /// Time period for date rules.
    pub time_period: Option<String>,
}

impl ConditionalRule {
    /// Create a cell value rule.
    pub fn cell_is(operator: ConditionalOperator, value: &str) -> Self {
        ConditionalRule {
            rule_type: ConditionalFormatType::CellIs,
            priority: 1,
            operator: Some(operator),
            formula1: Some(value.to_string()),
            formula2: None,
            text: None,
            format: None,
            color_scale: None,
            data_bar: None,
            icon_set: None,
            stop_if_true: false,
            bottom: false,
            percent: false,
            rank: None,
            above_average: true,
            equal_average: false,
            std_dev: None,
            time_period: None,
        }
    }

    /// Create a formula-based rule.
    pub fn formula(formula: &str) -> Self {
        ConditionalRule {
            rule_type: ConditionalFormatType::Expression,
            priority: 1,
            operator: None,
            formula1: Some(formula.to_string()),
            formula2: None,
            text: None,
            format: None,
            color_scale: None,
            data_bar: None,
            icon_set: None,
            stop_if_true: false,
            bottom: false,
            percent: false,
            rank: None,
            above_average: true,
            equal_average: false,
            std_dev: None,
            time_period: None,
        }
    }

    /// Create a color scale rule.
    pub fn with_color_scale(scale: ColorScale) -> Self {
        ConditionalRule {
            rule_type: ConditionalFormatType::ColorScale,
            priority: 1,
            operator: None,
            formula1: None,
            formula2: None,
            text: None,
            format: None,
            color_scale: Some(scale),
            data_bar: None,
            icon_set: None,
            stop_if_true: false,
            bottom: false,
            percent: false,
            rank: None,
            above_average: true,
            equal_average: false,
            std_dev: None,
            time_period: None,
        }
    }

    /// Create a data bar rule.
    pub fn with_data_bar(bar: DataBar) -> Self {
        ConditionalRule {
            rule_type: ConditionalFormatType::DataBar,
            priority: 1,
            operator: None,
            formula1: None,
            formula2: None,
            text: None,
            format: None,
            color_scale: None,
            data_bar: Some(bar),
            icon_set: None,
            stop_if_true: false,
            bottom: false,
            percent: false,
            rank: None,
            above_average: true,
            equal_average: false,
            std_dev: None,
            time_period: None,
        }
    }

    /// Create an icon set rule.
    pub fn with_icon_set(icons: IconSet) -> Self {
        ConditionalRule {
            rule_type: ConditionalFormatType::IconSet,
            priority: 1,
            operator: None,
            formula1: None,
            formula2: None,
            text: None,
            format: None,
            color_scale: None,
            data_bar: None,
            icon_set: Some(icons),
            stop_if_true: false,
            bottom: false,
            percent: false,
            rank: None,
            above_average: true,
            equal_average: false,
            std_dev: None,
            time_period: None,
        }
    }

    /// Create a top N rule.
    pub fn top(n: u32) -> Self {
        ConditionalRule {
            rule_type: ConditionalFormatType::Top10,
            priority: 1,
            operator: None,
            formula1: None,
            formula2: None,
            text: None,
            format: None,
            color_scale: None,
            data_bar: None,
            icon_set: None,
            stop_if_true: false,
            bottom: false,
            percent: false,
            rank: Some(n),
            above_average: true,
            equal_average: false,
            std_dev: None,
            time_period: None,
        }
    }

    /// Create a bottom N rule.
    pub fn bottom(n: u32) -> Self {
        let mut rule = Self::top(n);
        rule.bottom = true;
        rule
    }

    /// Set the format to apply.
    pub fn with_format(mut self, format: ConditionalFormat) -> Self {
        self.format = Some(format);
        self
    }

    /// Set the priority.
    pub fn with_priority(mut self, priority: u32) -> Self {
        self.priority = priority;
        self
    }
}

/// Conditional formatting for a range.
#[derive(Clone, Debug)]
pub struct ConditionalFormatting {
    /// Cell range (e.g., "A1:B10").
    pub range: String,
    /// Rules to apply.
    pub rules: Vec<ConditionalRule>,
}

impl ConditionalFormatting {
    /// Create conditional formatting for a range.
    pub fn new<S: Into<String>>(range: S) -> Self {
        ConditionalFormatting {
            range: range.into(),
            rules: Vec::new(),
        }
    }

    /// Add a rule.
    pub fn add_rule(&mut self, rule: ConditionalRule) {
        self.rules.push(rule);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_is_rule() {
        let rule = ConditionalRule::cell_is(ConditionalOperator::GreaterThan, "100")
            .with_format(ConditionalFormat::new().with_fill(ConditionalColor::green()));

        assert_eq!(rule.rule_type, ConditionalFormatType::CellIs);
        assert_eq!(rule.operator, Some(ConditionalOperator::GreaterThan));
        assert_eq!(rule.formula1, Some("100".to_string()));
    }

    #[test]
    fn test_color_scale() {
        let scale = ColorScale::red_yellow_green();
        let rule = ConditionalRule::with_color_scale(scale);

        assert_eq!(rule.rule_type, ConditionalFormatType::ColorScale);
        assert!(rule.color_scale.is_some());
    }

    #[test]
    fn test_data_bar() {
        let bar = DataBar::new().with_color(ConditionalColor::blue()).hide_value();
        let rule = ConditionalRule::with_data_bar(bar);

        assert_eq!(rule.rule_type, ConditionalFormatType::DataBar);
        assert!(rule.data_bar.is_some());
        assert!(!rule.data_bar.unwrap().show_value);
    }

    #[test]
    fn test_icon_set() {
        let icons = IconSet::new(IconSetStyle::ThreeArrows).icon_only();
        let rule = ConditionalRule::with_icon_set(icons);

        assert_eq!(rule.rule_type, ConditionalFormatType::IconSet);
        assert!(rule.icon_set.is_some());
    }

    #[test]
    fn test_conditional_formatting() {
        let mut cf = ConditionalFormatting::new("A1:A100");
        cf.add_rule(ConditionalRule::top(10).with_format(
            ConditionalFormat::new()
                .with_bold(true)
                .with_fill(ConditionalColor::yellow()),
        ));

        assert_eq!(cf.range, "A1:A100");
        assert_eq!(cf.rules.len(), 1);
    }
}
