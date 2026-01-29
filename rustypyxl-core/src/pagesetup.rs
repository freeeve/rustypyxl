//! Page setup and print settings for Excel worksheets.
//!
//! This module provides structures for configuring print layout, margins,
//! headers, footers, and other page setup options.

/// Paper size constants.
#[derive(Clone, Debug, PartialEq, Default)]
pub enum PaperSize {
    /// Letter (8.5" x 11")
    #[default]
    Letter,
    /// Legal (8.5" x 14")
    Legal,
    /// Executive (7.25" x 10.5")
    Executive,
    /// A3 (297mm x 420mm)
    A3,
    /// A4 (210mm x 297mm)
    A4,
    /// A5 (148mm x 210mm)
    A5,
    /// B4 (250mm x 353mm)
    B4,
    /// B5 (182mm x 257mm)
    B5,
    /// Tabloid (11" x 17")
    Tabloid,
    /// Custom size
    Custom(u32),
}

impl PaperSize {
    /// Get the numeric code for this paper size.
    pub fn code(&self) -> u32 {
        match self {
            PaperSize::Letter => 1,
            PaperSize::Legal => 5,
            PaperSize::Executive => 7,
            PaperSize::A3 => 8,
            PaperSize::A4 => 9,
            PaperSize::A5 => 11,
            PaperSize::B4 => 12,
            PaperSize::B5 => 13,
            PaperSize::Tabloid => 3,
            PaperSize::Custom(code) => *code,
        }
    }
}

/// Page orientation.
#[derive(Clone, Debug, PartialEq, Default)]
pub enum Orientation {
    /// Portrait (vertical)
    #[default]
    Portrait,
    /// Landscape (horizontal)
    Landscape,
}

/// Page order for printing.
#[derive(Clone, Debug, PartialEq, Default)]
pub enum PageOrder {
    /// Down then over (default)
    #[default]
    DownThenOver,
    /// Over then down
    OverThenDown,
}

/// Cell comments printing option.
#[derive(Clone, Debug, PartialEq, Default)]
pub enum CellComments {
    /// Don't print comments
    #[default]
    None,
    /// Print at end of sheet
    AtEnd,
    /// Print as displayed
    AsDisplayed,
}

/// Print errors option.
#[derive(Clone, Debug, PartialEq, Default)]
pub enum PrintErrors {
    /// Print as displayed
    #[default]
    Displayed,
    /// Print as blank
    Blank,
    /// Print as dashes
    Dash,
    /// Print as N/A
    NA,
}

/// Page margins in inches.
#[derive(Clone, Debug)]
pub struct PageMargins {
    /// Left margin
    pub left: f64,
    /// Right margin
    pub right: f64,
    /// Top margin
    pub top: f64,
    /// Bottom margin
    pub bottom: f64,
    /// Header margin
    pub header: f64,
    /// Footer margin
    pub footer: f64,
}

impl Default for PageMargins {
    fn default() -> Self {
        PageMargins {
            left: 0.7,
            right: 0.7,
            top: 0.75,
            bottom: 0.75,
            header: 0.3,
            footer: 0.3,
        }
    }
}

impl PageMargins {
    /// Create margins with all values the same.
    pub fn uniform(margin: f64) -> Self {
        PageMargins {
            left: margin,
            right: margin,
            top: margin,
            bottom: margin,
            header: margin / 2.0,
            footer: margin / 2.0,
        }
    }

    /// Narrow margins preset.
    pub fn narrow() -> Self {
        PageMargins {
            left: 0.25,
            right: 0.25,
            top: 0.75,
            bottom: 0.75,
            header: 0.3,
            footer: 0.3,
        }
    }

    /// Wide margins preset.
    pub fn wide() -> Self {
        PageMargins {
            left: 1.0,
            right: 1.0,
            top: 1.0,
            bottom: 1.0,
            header: 0.5,
            footer: 0.5,
        }
    }
}

/// Header or footer section.
#[derive(Clone, Debug, Default)]
pub struct HeaderFooterSection {
    /// Left section content
    pub left: Option<String>,
    /// Center section content
    pub center: Option<String>,
    /// Right section content
    pub right: Option<String>,
}

impl HeaderFooterSection {
    /// Create a new section.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the left content.
    pub fn with_left<S: Into<String>>(mut self, text: S) -> Self {
        self.left = Some(text.into());
        self
    }

    /// Set the center content.
    pub fn with_center<S: Into<String>>(mut self, text: S) -> Self {
        self.center = Some(text.into());
        self
    }

    /// Set the right content.
    pub fn with_right<S: Into<String>>(mut self, text: S) -> Self {
        self.right = Some(text.into());
        self
    }

    /// Build the header/footer string with section codes.
    pub fn to_string(&self) -> String {
        let mut result = String::new();
        if let Some(ref left) = self.left {
            result.push_str("&L");
            result.push_str(left);
        }
        if let Some(ref center) = self.center {
            result.push_str("&C");
            result.push_str(center);
        }
        if let Some(ref right) = self.right {
            result.push_str("&R");
            result.push_str(right);
        }
        result
    }
}

/// Header and footer configuration.
#[derive(Clone, Debug, Default)]
pub struct HeaderFooter {
    /// Odd page header
    pub odd_header: Option<HeaderFooterSection>,
    /// Odd page footer
    pub odd_footer: Option<HeaderFooterSection>,
    /// Even page header (when different)
    pub even_header: Option<HeaderFooterSection>,
    /// Even page footer (when different)
    pub even_footer: Option<HeaderFooterSection>,
    /// First page header (when different)
    pub first_header: Option<HeaderFooterSection>,
    /// First page footer (when different)
    pub first_footer: Option<HeaderFooterSection>,
    /// Different odd/even pages
    pub different_odd_even: bool,
    /// Different first page
    pub different_first: bool,
    /// Scale with document
    pub scale_with_doc: bool,
    /// Align with margins
    pub align_with_margins: bool,
}

impl HeaderFooter {
    /// Create new header/footer configuration.
    pub fn new() -> Self {
        HeaderFooter {
            scale_with_doc: true,
            align_with_margins: true,
            ..Default::default()
        }
    }

    /// Set the header for all pages.
    pub fn with_header(mut self, section: HeaderFooterSection) -> Self {
        self.odd_header = Some(section);
        self
    }

    /// Set the footer for all pages.
    pub fn with_footer(mut self, section: HeaderFooterSection) -> Self {
        self.odd_footer = Some(section);
        self
    }
}

/// Common header/footer codes.
pub mod codes {
    /// Current page number
    pub const PAGE_NUMBER: &str = "&P";
    /// Total pages
    pub const TOTAL_PAGES: &str = "&N";
    /// Current date
    pub const DATE: &str = "&D";
    /// Current time
    pub const TIME: &str = "&T";
    /// File path
    pub const FILE_PATH: &str = "&Z";
    /// File name
    pub const FILE_NAME: &str = "&F";
    /// Sheet name
    pub const SHEET_NAME: &str = "&A";
    /// Bold on/off
    pub const BOLD: &str = "&B";
    /// Italic on/off
    pub const ITALIC: &str = "&I";
    /// Underline on/off
    pub const UNDERLINE: &str = "&U";
    /// Strikethrough on/off
    pub const STRIKETHROUGH: &str = "&S";
    /// Font size (e.g., &12)
    pub fn font_size(size: u32) -> String {
        format!("&{}", size)
    }
    /// Font name (e.g., &"Arial")
    pub fn font_name(name: &str) -> String {
        format!("&\"{}\"", name)
    }
}

/// Print titles (rows/columns to repeat).
#[derive(Clone, Debug, Default)]
pub struct PrintTitles {
    /// Rows to repeat at top (e.g., "1:2")
    pub rows: Option<String>,
    /// Columns to repeat at left (e.g., "A:B")
    pub cols: Option<String>,
}

impl PrintTitles {
    /// Set rows to repeat.
    pub fn with_rows<S: Into<String>>(mut self, rows: S) -> Self {
        self.rows = Some(rows.into());
        self
    }

    /// Set columns to repeat.
    pub fn with_cols<S: Into<String>>(mut self, cols: S) -> Self {
        self.cols = Some(cols.into());
        self
    }
}

/// Page setup configuration.
#[derive(Clone, Debug)]
pub struct PageSetup {
    /// Paper size
    pub paper_size: PaperSize,
    /// Page orientation
    pub orientation: Orientation,
    /// Scale (percentage, 10-400)
    pub scale: u32,
    /// Fit to width (pages)
    pub fit_to_width: Option<u32>,
    /// Fit to height (pages)
    pub fit_to_height: Option<u32>,
    /// First page number
    pub first_page_number: Option<u32>,
    /// Page order
    pub page_order: PageOrder,
    /// Print black and white
    pub black_and_white: bool,
    /// Print draft quality
    pub draft: bool,
    /// Print cell comments
    pub cell_comments: CellComments,
    /// Print errors
    pub errors: PrintErrors,
    /// Horizontal DPI
    pub horizontal_dpi: Option<u32>,
    /// Vertical DPI
    pub vertical_dpi: Option<u32>,
    /// Number of copies
    pub copies: u32,
    /// Page margins
    pub margins: PageMargins,
    /// Header and footer
    pub header_footer: HeaderFooter,
    /// Print titles
    pub print_titles: PrintTitles,
    /// Print area
    pub print_area: Option<String>,
    /// Print gridlines
    pub print_gridlines: bool,
    /// Print row/column headings
    pub print_headings: bool,
    /// Center horizontally
    pub center_horizontally: bool,
    /// Center vertically
    pub center_vertically: bool,
}

impl Default for PageSetup {
    fn default() -> Self {
        PageSetup {
            paper_size: PaperSize::default(),
            orientation: Orientation::default(),
            scale: 100,
            fit_to_width: None,
            fit_to_height: None,
            first_page_number: None,
            page_order: PageOrder::default(),
            black_and_white: false,
            draft: false,
            cell_comments: CellComments::default(),
            errors: PrintErrors::default(),
            horizontal_dpi: None,
            vertical_dpi: None,
            copies: 1,
            margins: PageMargins::default(),
            header_footer: HeaderFooter::new(),
            print_titles: PrintTitles::default(),
            print_area: None,
            print_gridlines: false,
            print_headings: false,
            center_horizontally: false,
            center_vertically: false,
        }
    }
}

impl PageSetup {
    /// Create new page setup with default values.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set paper size.
    pub fn with_paper_size(mut self, size: PaperSize) -> Self {
        self.paper_size = size;
        self
    }

    /// Set orientation.
    pub fn with_orientation(mut self, orientation: Orientation) -> Self {
        self.orientation = orientation;
        self
    }

    /// Set scale percentage.
    pub fn with_scale(mut self, scale: u32) -> Self {
        self.scale = scale.clamp(10, 400);
        self
    }

    /// Fit to one page wide.
    pub fn fit_to_page(mut self) -> Self {
        self.fit_to_width = Some(1);
        self.fit_to_height = Some(1);
        self
    }

    /// Fit to width.
    pub fn fit_to_width(mut self, pages: u32) -> Self {
        self.fit_to_width = Some(pages);
        self
    }

    /// Set margins.
    pub fn with_margins(mut self, margins: PageMargins) -> Self {
        self.margins = margins;
        self
    }

    /// Set print area.
    pub fn with_print_area<S: Into<String>>(mut self, area: S) -> Self {
        self.print_area = Some(area.into());
        self
    }

    /// Enable gridline printing.
    pub fn print_gridlines(mut self) -> Self {
        self.print_gridlines = true;
        self
    }

    /// Center on page.
    pub fn center_on_page(mut self) -> Self {
        self.center_horizontally = true;
        self.center_vertically = true;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_page_setup_default() {
        let setup = PageSetup::new();
        assert_eq!(setup.paper_size, PaperSize::Letter);
        assert_eq!(setup.orientation, Orientation::Portrait);
        assert_eq!(setup.scale, 100);
    }

    #[test]
    fn test_page_setup_builder() {
        let setup = PageSetup::new()
            .with_paper_size(PaperSize::A4)
            .with_orientation(Orientation::Landscape)
            .with_scale(80)
            .print_gridlines()
            .center_on_page();

        assert_eq!(setup.paper_size, PaperSize::A4);
        assert_eq!(setup.orientation, Orientation::Landscape);
        assert_eq!(setup.scale, 80);
        assert!(setup.print_gridlines);
        assert!(setup.center_horizontally);
    }

    #[test]
    fn test_margins() {
        let margins = PageMargins::narrow();
        assert_eq!(margins.left, 0.25);
        assert_eq!(margins.right, 0.25);
    }

    #[test]
    fn test_header_footer_section() {
        let section = HeaderFooterSection::new()
            .with_left("Left text")
            .with_center(&format!("Page {} of {}", codes::PAGE_NUMBER, codes::TOTAL_PAGES))
            .with_right(codes::DATE);

        let s = section.to_string();
        assert!(s.contains("&L"));
        assert!(s.contains("&P"));
        assert!(s.contains("&D"));
    }

    #[test]
    fn test_paper_size_code() {
        assert_eq!(PaperSize::Letter.code(), 1);
        assert_eq!(PaperSize::A4.code(), 9);
    }
}
