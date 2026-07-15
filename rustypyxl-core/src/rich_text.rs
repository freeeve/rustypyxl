//! Rich text: a cell string whose runs carry their own formatting.
//!
//! In XLSX a string cell is normally a single `<t>`, but it can instead be a
//! sequence of `<r>` runs, each with an optional `<rPr>` (bold, italic, color,
//! ...) and its own `<t>`. Excel writes rich text into the shared-strings table.
//! rustypyxl preserves these runs so they survive a load->save round-trip; the
//! cell's plain `CellValue::String` remains the concatenated text.

use crate::style::Color;

/// The formatting of one rich-text run. `None` fields inherit the cell font.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct RunFont {
    pub bold: bool,
    pub italic: bool,
    /// Underline style: "single" or "double" (an empty `<u/>` means "single").
    pub underline: Option<String>,
    pub strike: bool,
    /// Point size.
    pub size: Option<f64>,
    pub color: Option<Color>,
    /// Font family name (`rFont`).
    pub name: Option<String>,
    /// "superscript" / "subscript".
    pub vert_align: Option<String>,
}

impl RunFont {
    /// True when no property is set (an `<rPr>` that would emit nothing).
    pub fn is_empty(&self) -> bool {
        !self.bold
            && !self.italic
            && self.underline.is_none()
            && !self.strike
            && self.size.is_none()
            && self.color.is_none()
            && self.name.is_none()
            && self.vert_align.is_none()
    }
}

/// One run of text with optional per-run formatting.
#[derive(Clone, Debug, PartialEq)]
pub struct TextRun {
    pub text: String,
    /// `None` = the run inherits the cell's font (a run with no `<rPr>`).
    pub font: Option<RunFont>,
}

impl TextRun {
    /// A run of unformatted text.
    pub fn plain<S: Into<String>>(text: S) -> Self {
        TextRun {
            text: text.into(),
            font: None,
        }
    }

    /// A run with the given font.
    pub fn formatted<S: Into<String>>(text: S, font: RunFont) -> Self {
        TextRun {
            text: text.into(),
            font: Some(font),
        }
    }
}

/// A rich-text string: an ordered list of runs.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct RichText {
    pub runs: Vec<TextRun>,
}

impl RichText {
    /// Build from runs.
    pub fn new(runs: Vec<TextRun>) -> Self {
        RichText { runs }
    }

    /// The concatenated plain text of every run.
    pub fn plain(&self) -> String {
        self.runs.iter().map(|r| r.text.as_str()).collect()
    }

    pub fn is_empty(&self) -> bool {
        self.runs.is_empty()
    }
}
