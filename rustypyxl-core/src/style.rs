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
    /// Strikethrough text.
    pub strike: bool,
    /// Font color as RGB hex (e.g., "#FF0000") or theme reference.
    pub color: Option<String>,
    /// Vertical alignment (superscript/subscript).
    pub vert_align: Option<String>,
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

    /// Set strikethrough style.
    pub fn with_strike(mut self, strike: bool) -> Self {
        self.strike = strike;
        self
    }

    /// Set font color.
    pub fn with_color<S: Into<String>>(mut self, color: S) -> Self {
        self.color = Some(color.into());
        self
    }

    /// Set vertical alignment (superscript/subscript).
    pub fn with_vert_align<S: Into<String>>(mut self, vert_align: S) -> Self {
        self.vert_align = Some(vert_align.into());
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

/// A color stop in a gradient fill.
#[derive(Clone, Debug, PartialEq)]
pub struct GradientStop {
    /// Position of the stop (0.0 to 1.0).
    pub position: f64,
    /// Color at this stop as RGB hex.
    pub color: String,
}

/// Gradient fill properties.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct GradientFill {
    /// Gradient type: "linear" or "path".
    pub gradient_type: Option<String>,
    /// Rotation angle in degrees for linear gradients.
    pub degree: Option<f64>,
    /// Left edge for path gradients.
    pub left: Option<f64>,
    /// Right edge for path gradients.
    pub right: Option<f64>,
    /// Top edge for path gradients.
    pub top: Option<f64>,
    /// Bottom edge for path gradients.
    pub bottom: Option<f64>,
    /// Color stops.
    pub stops: Vec<GradientStop>,
}

impl GradientFill {
    /// Create a new empty gradient fill.
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a linear gradient from two colors.
    pub fn linear<S1: Into<String>, S2: Into<String>>(start_color: S1, end_color: S2) -> Self {
        GradientFill {
            gradient_type: Some("linear".to_string()),
            degree: Some(90.0),
            left: None,
            right: None,
            top: None,
            bottom: None,
            stops: vec![
                GradientStop { position: 0.0, color: start_color.into() },
                GradientStop { position: 1.0, color: end_color.into() },
            ],
        }
    }

    /// Set the gradient type.
    pub fn with_type<S: Into<String>>(mut self, gradient_type: S) -> Self {
        self.gradient_type = Some(gradient_type.into());
        self
    }

    /// Set the rotation degree.
    pub fn with_degree(mut self, degree: f64) -> Self {
        self.degree = Some(degree);
        self
    }

    /// Add a color stop.
    pub fn with_stop<S: Into<String>>(mut self, position: f64, color: S) -> Self {
        self.stops.push(GradientStop { position, color: color.into() });
        self
    }
}

/// Cell protection properties.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Protection {
    /// Whether the cell is locked (default is true in Excel).
    pub locked: bool,
    /// Whether the formula is hidden when sheet is protected.
    pub hidden: bool,
}

impl Protection {
    /// Create a new Protection with default values (locked=true, hidden=false).
    pub fn new() -> Self {
        Protection {
            locked: true,
            hidden: false,
        }
    }

    /// Create an unlocked protection.
    pub fn unlocked() -> Self {
        Protection {
            locked: false,
            hidden: false,
        }
    }

    /// Set locked state.
    pub fn with_locked(mut self, locked: bool) -> Self {
        self.locked = locked;
        self
    }

    /// Set hidden state.
    pub fn with_hidden(mut self, hidden: bool) -> Self {
        self.hidden = hidden;
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
    /// Fill/background properties (pattern fill).
    pub fill: Option<Fill>,
    /// Gradient fill properties.
    pub gradient_fill: Option<GradientFill>,
    /// Number format string.
    pub number_format: Option<String>,
    /// Protection properties.
    pub protection: Option<Protection>,
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

    /// Set the gradient fill.
    pub fn with_gradient_fill(mut self, gradient_fill: GradientFill) -> Self {
        self.gradient_fill = Some(gradient_fill);
        self
    }

    /// Set the number format.
    pub fn with_number_format<S: Into<String>>(mut self, format: S) -> Self {
        self.number_format = Some(format.into());
        self
    }

    /// Set the protection.
    pub fn with_protection(mut self, protection: Protection) -> Self {
        self.protection = Some(protection);
        self
    }
}

/// A cell format entry (cellXf) that combines references to fonts, fills, borders, and number formats.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct CellXf {
    /// Index into the fonts array.
    pub font_id: usize,
    /// Index into the fills array.
    pub fill_id: usize,
    /// Index into the borders array.
    pub border_id: usize,
    /// Index into the number formats array (or built-in format ID).
    pub num_fmt_id: usize,
    /// Alignment properties (stored directly, not indexed).
    pub alignment: Option<Alignment>,
    /// Protection properties (stored directly, not indexed).
    pub protection: Option<Protection>,
    /// Whether font is applied.
    pub apply_font: bool,
    /// Whether fill is applied.
    pub apply_fill: bool,
    /// Whether border is applied.
    pub apply_border: bool,
    /// Whether number format is applied.
    pub apply_number_format: bool,
    /// Whether alignment is applied.
    pub apply_alignment: bool,
    /// Whether protection is applied.
    pub apply_protection: bool,
}

/// Registry of all styles in a workbook.
/// Excel stores styles as separate arrays of fonts, fills, borders, number formats,
/// and then cellXfs that combine them by index.
#[derive(Clone, Debug, Default)]
pub struct StyleRegistry {
    /// All fonts used in the workbook.
    pub fonts: Vec<Font>,
    /// All fills used in the workbook.
    pub fills: Vec<Fill>,
    /// All borders used in the workbook.
    pub borders: Vec<Border>,
    /// Custom number formats (format code -> format ID).
    pub num_fmts: Vec<(usize, String)>,
    /// Cell formats that combine font/fill/border/numFmt indices.
    pub cell_xfs: Vec<CellXf>,
}

impl StyleRegistry {
    /// Create a new empty style registry with Excel defaults.
    pub fn new() -> Self {
        let mut registry = StyleRegistry::default();

        // Excel requires at least one default font
        registry.fonts.push(Font {
            name: Some("Calibri".to_string()),
            size: Some(11.0),
            ..Default::default()
        });

        // Excel requires at least two fills (none and gray125)
        registry.fills.push(Fill::default()); // "none" pattern
        registry.fills.push(Fill {
            pattern_type: Some("gray125".to_string()),
            ..Default::default()
        });

        // Excel requires at least one border (empty)
        registry.borders.push(Border::default());

        // Default cell format (xf index 0)
        registry.cell_xfs.push(CellXf::default());

        registry
    }

    /// Get or create a font index.
    pub fn get_or_add_font(&mut self, font: &Font) -> usize {
        if let Some(idx) = self.fonts.iter().position(|f| f == font) {
            idx
        } else {
            let idx = self.fonts.len();
            self.fonts.push(font.clone());
            idx
        }
    }

    /// Get or create a fill index.
    pub fn get_or_add_fill(&mut self, fill: &Fill) -> usize {
        if let Some(idx) = self.fills.iter().position(|f| f == fill) {
            idx
        } else {
            let idx = self.fills.len();
            self.fills.push(fill.clone());
            idx
        }
    }

    /// Get or create a border index.
    pub fn get_or_add_border(&mut self, border: &Border) -> usize {
        if let Some(idx) = self.borders.iter().position(|b| b == border) {
            idx
        } else {
            let idx = self.borders.len();
            self.borders.push(border.clone());
            idx
        }
    }

    /// Get or create a number format ID.
    /// Built-in formats have IDs 0-163, custom formats start at 164.
    pub fn get_or_add_num_fmt(&mut self, format: &str) -> usize {
        // Check built-in formats first
        if let Some(id) = Self::builtin_num_fmt_id(format) {
            return id;
        }

        // Check existing custom formats
        if let Some((id, _)) = self.num_fmts.iter().find(|(_, f)| f == format) {
            return *id;
        }

        // Add new custom format (IDs start at 164)
        let id = 164 + self.num_fmts.len();
        self.num_fmts.push((id, format.to_string()));
        id
    }

    /// Get built-in number format ID for common formats.
    pub fn builtin_num_fmt_id(format: &str) -> Option<usize> {
        match format {
            "General" => Some(0),
            "0" => Some(1),
            "0.00" => Some(2),
            "#,##0" => Some(3),
            "#,##0.00" => Some(4),
            "0%" => Some(9),
            "0.00%" => Some(10),
            "0.00E+00" => Some(11),
            "mm-dd-yy" => Some(14),
            "d-mmm-yy" => Some(15),
            "d-mmm" => Some(16),
            "mmm-yy" => Some(17),
            "h:mm AM/PM" => Some(18),
            "h:mm:ss AM/PM" => Some(19),
            "h:mm" => Some(20),
            "h:mm:ss" => Some(21),
            "m/d/yy h:mm" => Some(22),
            "#,##0 ;(#,##0)" => Some(37),
            "#,##0 ;[Red](#,##0)" => Some(38),
            "#,##0.00;(#,##0.00)" => Some(39),
            "#,##0.00;[Red](#,##0.00)" => Some(40),
            "mm:ss" => Some(45),
            "[h]:mm:ss" => Some(46),
            "mmss.0" => Some(47),
            "##0.0E+0" => Some(48),
            "@" => Some(49),
            _ => None,
        }
    }

    /// Get or create a cell format (xf) index for a CellStyle.
    pub fn get_or_add_cell_xf(&mut self, style: &CellStyle) -> usize {
        let font_id = style.font.as_ref()
            .map(|f| self.get_or_add_font(f))
            .unwrap_or(0);

        let fill_id = style.fill.as_ref()
            .map(|f| self.get_or_add_fill(f))
            .unwrap_or(0);

        let border_id = style.border.as_ref()
            .map(|b| self.get_or_add_border(b))
            .unwrap_or(0);

        let num_fmt_id = style.number_format.as_ref()
            .map(|nf| self.get_or_add_num_fmt(nf))
            .unwrap_or(0);

        let xf = CellXf {
            font_id,
            fill_id,
            border_id,
            num_fmt_id,
            alignment: style.alignment.clone(),
            protection: style.protection.clone(),
            apply_font: style.font.is_some(),
            apply_fill: style.fill.is_some(),
            apply_border: style.border.is_some(),
            apply_number_format: style.number_format.is_some(),
            apply_alignment: style.alignment.is_some(),
            apply_protection: style.protection.is_some(),
        };

        // Check if this exact xf already exists
        if let Some(idx) = self.cell_xfs.iter().position(|x| x == &xf) {
            idx
        } else {
            let idx = self.cell_xfs.len();
            self.cell_xfs.push(xf);
            idx
        }
    }

    /// Build a CellStyle from a cell format index.
    pub fn get_cell_style(&self, xf_index: usize) -> Option<CellStyle> {
        let xf = self.cell_xfs.get(xf_index)?;

        let font = if xf.apply_font && xf.font_id < self.fonts.len() {
            Some(self.fonts[xf.font_id].clone())
        } else {
            None
        };

        let fill = if xf.apply_fill && xf.fill_id < self.fills.len() {
            Some(self.fills[xf.fill_id].clone())
        } else {
            None
        };

        let border = if xf.apply_border && xf.border_id < self.borders.len() {
            Some(self.borders[xf.border_id].clone())
        } else {
            None
        };

        let number_format = if xf.apply_number_format {
            self.get_num_fmt_string(xf.num_fmt_id)
        } else {
            None
        };

        let protection = if xf.apply_protection {
            xf.protection.clone()
        } else {
            None
        };

        Some(CellStyle {
            font,
            alignment: xf.alignment.clone(),
            border,
            fill,
            gradient_fill: None, // TODO: Add gradient fill support
            number_format,
            protection,
        })
    }

    /// Get the format string for a number format ID.
    fn get_num_fmt_string(&self, id: usize) -> Option<String> {
        // Check custom formats
        if let Some((_, fmt)) = self.num_fmts.iter().find(|(i, _)| *i == id) {
            return Some(fmt.clone());
        }

        // Return built-in format strings
        match id {
            0 => Some("General".to_string()),
            1 => Some("0".to_string()),
            2 => Some("0.00".to_string()),
            3 => Some("#,##0".to_string()),
            4 => Some("#,##0.00".to_string()),
            9 => Some("0%".to_string()),
            10 => Some("0.00%".to_string()),
            14 => Some("mm-dd-yy".to_string()),
            _ => None,
        }
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
