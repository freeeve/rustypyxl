//! Cell styling types: Font, Fill, Border, Alignment, CellStyle.

/// A color in a cell style.
///
/// Excel's `<color>` element is one of an explicit aRGB value, an index into
/// the workbook theme, or an index into the legacy palette -- any of which may
/// carry a `tint` that lightens or darkens it. Colors used to be stored as a
/// plain String with a `"theme:N"` sentinel, which could not express `tint` or
/// `indexed` at all, so both were dropped on load and on save.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Color {
    /// Explicit color as aRGB or RGB hex, with or without a leading '#'.
    pub rgb: Option<String>,
    /// Index into the workbook's theme color scheme.
    pub theme: Option<u32>,
    /// Index into the legacy indexed palette.
    pub indexed: Option<u32>,
    /// Tint applied to the color, -1.0 (darker) to 1.0 (lighter).
    pub tint: Option<f64>,
}

impl Color {
    /// A color from an explicit hex value.
    pub fn rgb<S: Into<String>>(rgb: S) -> Self {
        Color {
            rgb: Some(rgb.into()),
            ..Default::default()
        }
    }

    /// A color from a theme index.
    pub fn theme(theme: u32) -> Self {
        Color {
            theme: Some(theme),
            ..Default::default()
        }
    }

    /// A color from the legacy indexed palette.
    pub fn indexed(indexed: u32) -> Self {
        Color {
            indexed: Some(indexed),
            ..Default::default()
        }
    }

    /// Apply a tint, -1.0 (darker) to 1.0 (lighter).
    pub fn with_tint(mut self, tint: f64) -> Self {
        self.tint = Some(tint);
        self
    }

    /// True when nothing is set, i.e. there is no color at all.
    pub fn is_empty(&self) -> bool {
        self.rgb.is_none() && self.theme.is_none() && self.indexed.is_none()
    }

    /// The hex value with any leading '#' removed and an alpha channel, which
    /// is the form the `rgb` XML attribute takes.
    pub fn argb(&self) -> Option<String> {
        let hex = self.rgb.as_deref()?;
        let hex = hex.strip_prefix('#').unwrap_or(hex);
        Some(if hex.len() >= 8 {
            hex.to_string()
        } else {
            format!("FF{}", hex)
        })
    }
}

/// Accepts the plain hex strings the API has always taken, plus the legacy
/// `"theme:N"` form that colors used to be stored as.
impl<S: AsRef<str>> From<S> for Color {
    fn from(value: S) -> Self {
        let value = value.as_ref();
        if let Some(theme) = value.strip_prefix("theme:") {
            if let Ok(theme) = theme.parse() {
                return Color::theme(theme);
            }
        }
        if let Some(indexed) = value.strip_prefix("indexed:") {
            if let Ok(indexed) = indexed.parse() {
                return Color::indexed(indexed);
            }
        }
        Color::rgb(value)
    }
}

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
    /// Underline style: `None`, or one of "single", "double", "singleAccounting",
    /// "doubleAccounting". An empty `<u/>` element is treated as "single".
    pub underline: Option<String>,
    /// Strikethrough text.
    pub strike: bool,
    /// Font color.
    pub color: Option<Color>,
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

    /// Set underline style (e.g. "single" or "double").
    pub fn with_underline<S: Into<String>>(mut self, underline: S) -> Self {
        self.underline = Some(underline.into());
        self
    }

    /// Set strikethrough style.
    pub fn with_strike(mut self, strike: bool) -> Self {
        self.strike = strike;
        self
    }

    /// Set font color.
    pub fn with_color<C: Into<Color>>(mut self, color: C) -> Self {
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
    /// Border color.
    pub color: Option<Color>,
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
    pub fn with_color<C: Into<Color>>(mut self, color: C) -> Self {
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
    /// Foreground color.
    pub fg_color: Option<Color>,
    /// Background color.
    pub bg_color: Option<Color>,
}

impl Fill {
    /// Create a new Fill with default values.
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a solid fill with the specified color.
    pub fn solid<C: Into<Color>>(color: C) -> Self {
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
    pub fn with_fg_color<C: Into<Color>>(mut self, color: C) -> Self {
        self.fg_color = Some(color.into());
        self
    }

    /// Set the background color.
    pub fn with_bg_color<C: Into<Color>>(mut self, color: C) -> Self {
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
                GradientStop {
                    position: 0.0,
                    color: start_color.into(),
                },
                GradientStop {
                    position: 1.0,
                    color: end_color.into(),
                },
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
        self.stops.push(GradientStop {
            position,
            color: color.into(),
        });
        self
    }
}

/// Cell protection properties.
#[derive(Clone, Debug, PartialEq)]
pub struct Protection {
    /// Whether the cell is locked (default is true in Excel).
    pub locked: bool,
    /// Whether the formula is hidden when sheet is protected.
    pub hidden: bool,
}

/// Default matches Excel semantics (locked=true); a derived Default would
/// silently unlock cells.
impl Default for Protection {
    fn default() -> Self {
        Protection {
            locked: true,
            hidden: false,
        }
    }
}

impl Protection {
    /// Create a new Protection with default values (locked=true, hidden=false).
    pub fn new() -> Self {
        Self::default()
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
    /// Number format string. Interned so that cloning a style -- which happens
    /// per cell while resolving styles on save -- is a refcount bump.
    pub number_format: Option<crate::cell::InternedString>,
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
    pub fn with_number_format<S: AsRef<str>>(mut self, format: S) -> Self {
        self.number_format = Some(std::sync::Arc::from(format.as_ref()));
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
    /// Differential formats parsed from styles.xml `<dxfs>`, indexed by
    /// dxfId. Only populated on load; save regenerates the list from the
    /// conditional-formatting rules themselves.
    pub dxfs: Vec<crate::conditional::ConditionalFormat>,
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

        // Add new custom format. IDs start at 164; allocate past the highest
        // existing id rather than 164 + len(), which collides with the
        // non-contiguous ids found in loaded files.
        let id = self
            .num_fmts
            .iter()
            .map(|(id, _)| *id + 1)
            .max()
            .unwrap_or(164)
            .max(164);
        self.num_fmts.push((id, format.to_string()));
        id
    }

    /// Format code for a built-in number format id, the reverse of
    /// builtin_num_fmt_id. Built-ins are implicit in styles.xml, so loaders
    /// must resolve the id themselves.
    pub fn builtin_num_fmt_code(id: u32) -> Option<&'static str> {
        Some(match id {
            0 => "General",
            1 => "0",
            2 => "0.00",
            3 => "#,##0",
            4 => "#,##0.00",
            9 => "0%",
            10 => "0.00%",
            11 => "0.00E+00",
            12 => "# ?/?",
            13 => "# ??/??",
            14 => "mm-dd-yy",
            15 => "d-mmm-yy",
            16 => "d-mmm",
            17 => "mmm-yy",
            18 => "h:mm AM/PM",
            19 => "h:mm:ss AM/PM",
            20 => "h:mm",
            21 => "h:mm:ss",
            22 => "m/d/yy h:mm",
            37 => "#,##0 ;(#,##0)",
            38 => "#,##0 ;[Red](#,##0)",
            39 => "#,##0.00;(#,##0.00)",
            40 => "#,##0.00;[Red](#,##0.00)",
            45 => "mm:ss",
            46 => "[h]:mm:ss",
            47 => "mmss.0",
            48 => "##0.0E+0",
            49 => "@",
            _ => return None,
        })
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
        let font_id = style
            .font
            .as_ref()
            .map(|f| self.get_or_add_font(f))
            .unwrap_or(0);

        let fill_id = style
            .fill
            .as_ref()
            .map(|f| self.get_or_add_fill(f))
            .unwrap_or(0);

        let border_id = style
            .border
            .as_ref()
            .map(|b| self.get_or_add_border(b))
            .unwrap_or(0);

        let num_fmt_id = style
            .number_format
            .as_ref()
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
    fn get_num_fmt_string(&self, id: usize) -> Option<crate::cell::InternedString> {
        // Check custom formats
        if let Some((_, fmt)) = self.num_fmts.iter().find(|(i, _)| *i == id) {
            return Some(std::sync::Arc::from(fmt.as_str()));
        }

        // Built-ins are implicit in the file but must round-trip in the model
        Self::builtin_num_fmt_code(id as u32).map(std::sync::Arc::from)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_custom_num_fmt_ids_skip_loaded_ids() {
        let mut reg = StyleRegistry::new();
        // Simulate a loaded file carrying a non-contiguous custom id
        reg.num_fmts.push((165, "yyyy".to_string()));
        // A new format must not reuse 165 (the old 164+len scheme did)
        assert_eq!(reg.get_or_add_num_fmt("0.000%"), 166);
        // An existing format keeps its id
        assert_eq!(reg.get_or_add_num_fmt("yyyy"), 165);
        // Built-ins resolve to their fixed ids
        assert_eq!(reg.get_or_add_num_fmt("0.00%"), 10);
    }

    #[test]
    fn test_builtin_num_fmt_tables_are_inverse() {
        for id in [0u32, 1, 2, 3, 4, 9, 10, 11, 14, 20, 21, 22, 45, 49] {
            let code = StyleRegistry::builtin_num_fmt_code(id).unwrap();
            assert_eq!(
                StyleRegistry::builtin_num_fmt_id(code),
                Some(id as usize),
                "{}",
                code
            );
        }
    }

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
        assert_eq!(font.color, Some(Color::rgb("#FF0000")));
    }

    /// A Color knows which of the three ways it identifies itself, and tint
    /// rides along with any of them.
    #[test]
    fn test_color_kinds() {
        assert_eq!(Color::from("FF0000"), Color::rgb("FF0000"));
        assert_eq!(Color::from("theme:3"), Color::theme(3));
        assert_eq!(Color::from("indexed:9"), Color::indexed(9));

        let tinted = Color::theme(1).with_tint(-0.5);
        assert_eq!(tinted.theme, Some(1));
        assert_eq!(tinted.tint, Some(-0.5));
        assert!(!tinted.is_empty());
        assert!(Color::default().is_empty());

        // argb() supplies the alpha channel the rgb attribute needs
        assert_eq!(Color::rgb("FF0000").argb().as_deref(), Some("FFFF0000"));
        assert_eq!(Color::rgb("#00FF00").argb().as_deref(), Some("FF00FF00"));
        assert_eq!(Color::rgb("8000FF00").argb().as_deref(), Some("8000FF00"));
        assert_eq!(Color::theme(1).argb(), None);
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
        assert_eq!(fill.fg_color, Some(Color::rgb("#FFFF00")));
    }

    #[test]
    fn test_cell_style_builder() {
        let style = CellStyle::new()
            .with_font(Font::new().with_bold(true))
            .with_alignment(Alignment::new().with_horizontal("center"))
            .with_number_format("#,##0.00");

        assert!(style.font.is_some());
        assert!(style.alignment.is_some());
        assert_eq!(style.number_format.as_deref(), Some("#,##0.00"));
    }
}

#[cfg(test)]
mod coverage_tests {
    use super::*;

    #[test]
    fn color_constructors_and_argb() {
        assert_eq!(Color::rgb("FF0000").rgb.as_deref(), Some("FF0000"));
        assert_eq!(Color::theme(4).theme, Some(4));
        assert_eq!(Color::indexed(64).indexed, Some(64));
        assert_eq!(Color::rgb("00FF00").with_tint(0.5).tint, Some(0.5));

        assert!(Color {
            rgb: None,
            theme: None,
            indexed: None,
            tint: None
        }
        .is_empty());
        assert!(!Color::rgb("000000").is_empty());

        // argb pads a 6-digit hex with an alpha channel and strips '#'
        assert_eq!(Color::rgb("#FF0000").argb().as_deref(), Some("FFFF0000"));
        assert_eq!(Color::rgb("FF00FF00").argb().as_deref(), Some("FF00FF00"));
        assert_eq!(Color::theme(1).argb(), None);
    }

    #[test]
    fn color_from_str_forms() {
        assert_eq!(Color::from("theme:3").theme, Some(3));
        assert_eq!(Color::from("indexed:9").indexed, Some(9));
        assert_eq!(Color::from("ABCDEF").rgb.as_deref(), Some("ABCDEF"));
    }

    #[test]
    fn font_builders() {
        let f = Font::new()
            .with_name("Arial")
            .with_size(12.0)
            .with_bold(true)
            .with_italic(true)
            .with_underline("single")
            .with_strike(true)
            .with_color("FF0000")
            .with_vert_align("superscript");
        assert_eq!(f.name.as_deref(), Some("Arial"));
        assert_eq!(f.size, Some(12.0));
        assert!(f.bold && f.italic && f.strike);
        assert_eq!(f.underline.as_deref(), Some("single"));
        assert_eq!(f.color.unwrap().rgb.as_deref(), Some("FF0000"));
        assert_eq!(f.vert_align.as_deref(), Some("superscript"));
    }

    #[test]
    fn alignment_builders() {
        let a = Alignment::new()
            .with_horizontal("center")
            .with_vertical("top")
            .with_wrap_text(true);
        assert_eq!(a.horizontal.as_deref(), Some("center"));
        assert_eq!(a.vertical.as_deref(), Some("top"));
        assert!(a.wrap_text);
    }

    #[test]
    fn border_builders() {
        assert_eq!(BorderStyle::thin().style, "thin");
        assert_eq!(BorderStyle::medium().style, "medium");
        assert_eq!(BorderStyle::thick().style, "thick");
        assert_eq!(
            BorderStyle::new("dashed")
                .with_color("000000")
                .color
                .unwrap()
                .rgb
                .as_deref(),
            Some("000000")
        );

        let all = Border::all(BorderStyle::thin());
        assert!(
            all.left.is_some() && all.right.is_some() && all.top.is_some() && all.bottom.is_some()
        );

        let b = Border::new()
            .with_left(BorderStyle::thin())
            .with_right(BorderStyle::medium())
            .with_top(BorderStyle::thick())
            .with_bottom(BorderStyle::thin());
        assert_eq!(b.left.unwrap().style, "thin");
        assert_eq!(b.right.unwrap().style, "medium");
        assert_eq!(b.top.unwrap().style, "thick");
        assert_eq!(b.bottom.unwrap().style, "thin");
    }

    #[test]
    fn fill_builders() {
        let solid = Fill::solid("FFFF00");
        assert_eq!(solid.pattern_type.as_deref(), Some("solid"));
        assert_eq!(solid.fg_color.unwrap().rgb.as_deref(), Some("FFFF00"));

        let f = Fill::new()
            .with_pattern("gray125")
            .with_fg_color("111111")
            .with_bg_color("222222");
        assert_eq!(f.pattern_type.as_deref(), Some("gray125"));
        assert_eq!(f.fg_color.unwrap().rgb.as_deref(), Some("111111"));
        assert_eq!(f.bg_color.unwrap().rgb.as_deref(), Some("222222"));
    }

    #[test]
    fn gradient_fill_builders() {
        let g = GradientFill::linear("FF0000", "0000FF")
            .with_type("linear")
            .with_degree(90.0)
            .with_stop(0.5, "00FF00");
        assert_eq!(g.gradient_type.as_deref(), Some("linear"));
        assert_eq!(g.degree, Some(90.0));
        // linear() seeds two stops; with_stop adds a third
        assert_eq!(g.stops.len(), 3);
    }

    #[test]
    fn protection_builders() {
        assert!(Protection::new().locked);
        assert!(!Protection::unlocked().locked);
        let p = Protection::new().with_locked(false).with_hidden(true);
        assert!(!p.locked && p.hidden);
    }

    #[test]
    fn cell_style_builders() {
        let s = CellStyle::new()
            .with_font(Font::new().with_bold(true))
            .with_alignment(Alignment::new().with_wrap_text(true))
            .with_border(Border::all(BorderStyle::thin()))
            .with_fill(Fill::solid("EEEEEE"))
            .with_gradient_fill(GradientFill::linear("FFF", "000"))
            .with_number_format("0.00%")
            .with_protection(Protection::unlocked());
        assert!(s.font.is_some() && s.alignment.is_some() && s.border.is_some());
        assert!(s.fill.is_some() && s.gradient_fill.is_some() && s.protection.is_some());
        assert_eq!(s.number_format.as_deref(), Some("0.00%"));
    }

    #[test]
    fn style_registry_dedup_and_num_fmt() {
        let mut reg = StyleRegistry::new();
        let font = Font::new().with_bold(true);
        let a = reg.get_or_add_font(&font);
        let b = reg.get_or_add_font(&font);
        assert_eq!(a, b, "identical fonts dedup to one index");

        let f1 = reg.get_or_add_fill(&Fill::solid("FF0000"));
        assert_eq!(f1, reg.get_or_add_fill(&Fill::solid("FF0000")));
        let bd = reg.get_or_add_border(&Border::all(BorderStyle::thin()));
        assert_eq!(bd, reg.get_or_add_border(&Border::all(BorderStyle::thin())));

        // Built-in formats resolve to their id; custom formats allocate >= 164.
        assert_eq!(reg.get_or_add_num_fmt("0.00"), 2);
        let custom = reg.get_or_add_num_fmt("0.000\"x\"");
        assert!(custom >= 164);
        assert_eq!(custom, reg.get_or_add_num_fmt("0.000\"x\""));
    }

    #[test]
    fn builtin_num_fmt_roundtrip() {
        assert_eq!(StyleRegistry::builtin_num_fmt_code(2), Some("0.00"));
        assert_eq!(StyleRegistry::builtin_num_fmt_code(49), Some("@"));
        assert_eq!(StyleRegistry::builtin_num_fmt_code(200), None);
        assert_eq!(StyleRegistry::builtin_num_fmt_id("0.00"), Some(2));
        assert_eq!(StyleRegistry::builtin_num_fmt_id("General"), Some(0));
    }

    #[test]
    fn cell_xf_roundtrip_through_registry() {
        let mut reg = StyleRegistry::new();
        let style = CellStyle::new()
            .with_font(Font::new().with_size(14.0))
            .with_number_format("0.00%");
        let xf = reg.get_or_add_cell_xf(&style);
        let recovered = reg.get_cell_style(xf).expect("xf resolves back to a style");
        assert_eq!(recovered.number_format.as_deref(), Some("0.00%"));
        assert_eq!(recovered.font.and_then(|f| f.size), Some(14.0));
    }
}
