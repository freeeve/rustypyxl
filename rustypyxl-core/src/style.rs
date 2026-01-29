//! Cell styling types: Font, Fill, Border, Alignment, CellStyle.

/// Font properties for cell styling.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Font {
    /// Font family name (e.g., "Calibri", "Arial").
    pub name: Option<String>,
    /// Font size in points.
    pub size: Option<f64>,
    /// Bold text.
    pub bold: bool,
    /// Italic text.
    pub italic: bool,
    /// Underline text.
    pub underline: bool,
    /// Font color as RGB hex (e.g., "#FF0000") or theme reference.
    pub color: Option<String>,
}

impl Font {
    /// Create a new Font with default values.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the font name.
    pub fn with_name<S: Into<String>>(mut self, name: S) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Set the font size.
    pub fn with_size(mut self, size: f64) -> Self {
        self.size = Some(size);
        self
    }

    /// Set bold style.
    pub fn with_bold(mut self, bold: bool) -> Self {
        self.bold = bold;
        self
    }

    /// Set italic style.
    pub fn with_italic(mut self, italic: bool) -> Self {
        self.italic = italic;
        self
    }

    /// Set underline style.
    pub fn with_underline(mut self, underline: bool) -> Self {
        self.underline = underline;
        self
    }

    /// Set font color.
    pub fn with_color<S: Into<String>>(mut self, color: S) -> Self {
        self.color = Some(color.into());
        self
    }
}

/// Text alignment properties.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Alignment {
    /// Horizontal alignment: left, center, right, fill, justify, etc.
    pub horizontal: Option<String>,
    /// Vertical alignment: top, center, bottom, justify, distributed.
    pub vertical: Option<String>,
    /// Wrap text within cell.
    pub wrap_text: bool,
    /// Text rotation angle (-90 to 90).
    pub text_rotation: Option<i32>,
    /// Indent level.
    pub indent: Option<u32>,
    /// Shrink text to fit cell.
    pub shrink_to_fit: bool,
}

impl Alignment {
    /// Create a new Alignment with default values.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set horizontal alignment.
    pub fn with_horizontal<S: Into<String>>(mut self, align: S) -> Self {
        self.horizontal = Some(align.into());
        self
    }

    /// Set vertical alignment.
    pub fn with_vertical<S: Into<String>>(mut self, align: S) -> Self {
        self.vertical = Some(align.into());
        self
    }

    /// Set wrap text.
    pub fn with_wrap_text(mut self, wrap: bool) -> Self {
        self.wrap_text = wrap;
        self
    }
}

/// Border style for a single edge.
#[derive(Clone, Debug, PartialEq)]
pub struct BorderStyle {
    /// Border style: thin, medium, thick, dashed, dotted, double, etc.
    pub style: String,
    /// Border color as RGB hex.
    pub color: Option<String>,
}

impl BorderStyle {
    /// Create a new border style.
    pub fn new<S: Into<String>>(style: S) -> Self {
        BorderStyle {
            style: style.into(),
            color: None,
        }
    }

    /// Create a thin border.
    pub fn thin() -> Self {
        Self::new("thin")
    }

    /// Create a medium border.
    pub fn medium() -> Self {
        Self::new("medium")
    }

    /// Create a thick border.
    pub fn thick() -> Self {
        Self::new("thick")
    }

    /// Set the border color.
    pub fn with_color<S: Into<String>>(mut self, color: S) -> Self {
        self.color = Some(color.into());
        self
    }
}

/// Cell border properties.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Border {
    /// Left border.
    pub left: Option<BorderStyle>,
    /// Right border.
    pub right: Option<BorderStyle>,
    /// Top border.
    pub top: Option<BorderStyle>,
    /// Bottom border.
    pub bottom: Option<BorderStyle>,
    /// Diagonal border.
    pub diagonal: Option<BorderStyle>,
}

impl Border {
    /// Create a new Border with no edges.
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a border with all edges the same style.
    pub fn all(style: BorderStyle) -> Self {
        Border {
            left: Some(style.clone()),
            right: Some(style.clone()),
            top: Some(style.clone()),
            bottom: Some(style),
            diagonal: None,
        }
    }

    /// Set left border.
    pub fn with_left(mut self, style: BorderStyle) -> Self {
        self.left = Some(style);
        self
    }

    /// Set right border.
    pub fn with_right(mut self, style: BorderStyle) -> Self {
        self.right = Some(style);
        self
    }

    /// Set top border.
    pub fn with_top(mut self, style: BorderStyle) -> Self {
        self.top = Some(style);
        self
    }

    /// Set bottom border.
    pub fn with_bottom(mut self, style: BorderStyle) -> Self {
        self.bottom = Some(style);
        self
    }
}

/// Cell fill/background properties.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Fill {
    /// Pattern type: solid, gray125, darkGray, etc.
    pub pattern_type: Option<String>,
    /// Foreground color as RGB hex.
    pub fg_color: Option<String>,
    /// Background color as RGB hex.
    pub bg_color: Option<String>,
}

impl Fill {
    /// Create a new Fill with default values.
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a solid fill with the specified color.
    pub fn solid<S: Into<String>>(color: S) -> Self {
        Fill {
            pattern_type: Some("solid".to_string()),
            fg_color: Some(color.into()),
            bg_color: None,
        }
    }

    /// Set the pattern type.
    pub fn with_pattern<S: Into<String>>(mut self, pattern: S) -> Self {
        self.pattern_type = Some(pattern.into());
        self
    }

    /// Set the foreground color.
    pub fn with_fg_color<S: Into<String>>(mut self, color: S) -> Self {
        self.fg_color = Some(color.into());
        self
    }

    /// Set the background color.
    pub fn with_bg_color<S: Into<String>>(mut self, color: S) -> Self {
        self.bg_color = Some(color.into());
        self
    }
}

/// Complete cell style combining all styling components.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct CellStyle {
    /// Font properties.
    pub font: Option<Font>,
    /// Alignment properties.
    pub alignment: Option<Alignment>,
    /// Border properties.
    pub border: Option<Border>,
    /// Fill/background properties.
    pub fill: Option<Fill>,
    /// Number format string.
    pub number_format: Option<String>,
}

impl CellStyle {
    /// Create a new empty cell style.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the font.
    pub fn with_font(mut self, font: Font) -> Self {
        self.font = Some(font);
        self
    }

    /// Set the alignment.
    pub fn with_alignment(mut self, alignment: Alignment) -> Self {
        self.alignment = Some(alignment);
        self
    }

    /// Set the border.
    pub fn with_border(mut self, border: Border) -> Self {
        self.border = Some(border);
        self
    }

    /// Set the fill.
    pub fn with_fill(mut self, fill: Fill) -> Self {
        self.fill = Some(fill);
        self
    }

    /// Set the number format.
    pub fn with_number_format<S: Into<String>>(mut self, format: S) -> Self {
        self.number_format = Some(format.into());
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_font_builder() {
        let font = Font::new()
            .with_name("Arial")
            .with_size(12.0)
            .with_bold(true)
            .with_color("#FF0000");

        assert_eq!(font.name, Some("Arial".to_string()));
        assert_eq!(font.size, Some(12.0));
        assert!(font.bold);
        assert_eq!(font.color, Some("#FF0000".to_string()));
    }

    #[test]
    fn test_alignment_builder() {
        let align = Alignment::new()
            .with_horizontal("center")
            .with_vertical("top")
            .with_wrap_text(true);

        assert_eq!(align.horizontal, Some("center".to_string()));
        assert_eq!(align.vertical, Some("top".to_string()));
        assert!(align.wrap_text);
    }

    #[test]
    fn test_border_all() {
        let border = Border::all(BorderStyle::thin());
        assert!(border.left.is_some());
        assert!(border.right.is_some());
        assert!(border.top.is_some());
        assert!(border.bottom.is_some());
    }

    #[test]
    fn test_fill_solid() {
        let fill = Fill::solid("#FFFF00");
        assert_eq!(fill.pattern_type, Some("solid".to_string()));
        assert_eq!(fill.fg_color, Some("#FFFF00".to_string()));
    }

    #[test]
    fn test_cell_style_builder() {
        let style = CellStyle::new()
            .with_font(Font::new().with_bold(true))
            .with_alignment(Alignment::new().with_horizontal("center"))
            .with_number_format("#,##0.00");

        assert!(style.font.is_some());
        assert!(style.alignment.is_some());
        assert_eq!(style.number_format, Some("#,##0.00".to_string()));
    }
}
