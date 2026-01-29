//! Python bindings for styling classes.

#![allow(non_snake_case)]

use pyo3::prelude::*;

/// Font styling (openpyxl-compatible).
#[pyclass(name = "Font")]
#[derive(Clone, Debug, Default)]
pub struct PyFont {
    #[pyo3(get, set)]
    pub name: Option<String>,
    #[pyo3(get, set)]
    pub size: Option<f64>,
    #[pyo3(get, set)]
    pub bold: bool,
    #[pyo3(get, set)]
    pub italic: bool,
    #[pyo3(get, set)]
    pub underline: Option<String>,
    #[pyo3(get, set)]
    pub strike: bool,
    #[pyo3(get, set)]
    pub color: Option<String>,
    #[pyo3(get, set)]
    pub vertAlign: Option<String>,
}

#[pymethods]
impl PyFont {
    #[new]
    #[pyo3(signature = (name=None, size=None, bold=false, italic=false, underline=None, strike=false, color=None, vertAlign=None))]
    fn new(
        name: Option<String>,
        size: Option<f64>,
        bold: bool,
        italic: bool,
        underline: Option<String>,
        strike: bool,
        color: Option<String>,
        vertAlign: Option<String>,
    ) -> Self {
        PyFont {
            name,
            size,
            bold,
            italic,
            underline,
            strike,
            color,
            vertAlign,
        }
    }

    fn copy(&self) -> PyFont {
        self.clone()
    }

    fn __str__(&self) -> String {
        format!(
            "<Font name={:?} size={:?} bold={} italic={}>",
            self.name, self.size, self.bold, self.italic
        )
    }

    fn __repr__(&self) -> String {
        self.__str__()
    }
}

/// Text alignment (openpyxl-compatible).
#[pyclass(name = "Alignment")]
#[derive(Clone, Debug, Default)]
pub struct PyAlignment {
    #[pyo3(get, set)]
    pub horizontal: Option<String>,
    #[pyo3(get, set)]
    pub vertical: Option<String>,
    #[pyo3(get, set)]
    pub wrap_text: bool,
    #[pyo3(get, set)]
    pub shrink_to_fit: bool,
    #[pyo3(get, set)]
    pub indent: u32,
    #[pyo3(get, set)]
    pub text_rotation: i32,
}

#[pymethods]
impl PyAlignment {
    #[new]
    #[pyo3(signature = (horizontal=None, vertical=None, wrap_text=false, shrink_to_fit=false, indent=0, text_rotation=0))]
    fn new(
        horizontal: Option<String>,
        vertical: Option<String>,
        wrap_text: bool,
        shrink_to_fit: bool,
        indent: u32,
        text_rotation: i32,
    ) -> Self {
        PyAlignment {
            horizontal,
            vertical,
            wrap_text,
            shrink_to_fit,
            indent,
            text_rotation,
        }
    }

    fn copy(&self) -> PyAlignment {
        self.clone()
    }

    fn __str__(&self) -> String {
        format!(
            "<Alignment horizontal={:?} vertical={:?} wrap_text={}>",
            self.horizontal, self.vertical, self.wrap_text
        )
    }

    fn __repr__(&self) -> String {
        self.__str__()
    }
}

/// Pattern fill (openpyxl-compatible).
#[pyclass(name = "PatternFill")]
#[derive(Clone, Debug, Default)]
pub struct PyPatternFill {
    #[pyo3(get, set)]
    pub fill_type: Option<String>,
    #[pyo3(get, set)]
    pub fgColor: Option<String>,
    #[pyo3(get, set)]
    pub bgColor: Option<String>,
    #[pyo3(get, set)]
    pub patternType: Option<String>,
}

#[pymethods]
impl PyPatternFill {
    #[new]
    #[pyo3(signature = (fill_type=None, fgColor=None, bgColor=None, patternType=None))]
    fn new(
        fill_type: Option<String>,
        fgColor: Option<String>,
        bgColor: Option<String>,
        patternType: Option<String>,
    ) -> Self {
        PyPatternFill {
            fill_type: fill_type.or(patternType.clone()),
            fgColor,
            bgColor,
            patternType,
        }
    }

    fn copy(&self) -> PyPatternFill {
        self.clone()
    }

    fn __str__(&self) -> String {
        format!(
            "<PatternFill fill_type={:?} fgColor={:?}>",
            self.fill_type, self.fgColor
        )
    }

    fn __repr__(&self) -> String {
        self.__str__()
    }
}

/// Border style for a single edge (openpyxl-compatible).
#[pyclass(name = "Side")]
#[derive(Clone, Debug, Default)]
pub struct PySide {
    #[pyo3(get, set)]
    pub style: Option<String>,
    #[pyo3(get, set)]
    pub color: Option<String>,
}

#[pymethods]
impl PySide {
    #[new]
    #[pyo3(signature = (style=None, color=None))]
    fn new(style: Option<String>, color: Option<String>) -> Self {
        PySide { style, color }
    }

    fn copy(&self) -> PySide {
        self.clone()
    }

    fn __str__(&self) -> String {
        format!("<Side style={:?} color={:?}>", self.style, self.color)
    }

    fn __repr__(&self) -> String {
        self.__str__()
    }
}

/// Border (openpyxl-compatible).
#[pyclass(name = "Border")]
#[derive(Clone, Debug, Default)]
pub struct PyBorder {
    #[pyo3(get, set)]
    pub left: Option<PySide>,
    #[pyo3(get, set)]
    pub right: Option<PySide>,
    #[pyo3(get, set)]
    pub top: Option<PySide>,
    #[pyo3(get, set)]
    pub bottom: Option<PySide>,
    #[pyo3(get, set)]
    pub diagonal: Option<PySide>,
    #[pyo3(get, set)]
    pub diagonal_direction: Option<String>,
    #[pyo3(get, set)]
    pub outline: bool,
}

#[pymethods]
impl PyBorder {
    #[new]
    #[pyo3(signature = (left=None, right=None, top=None, bottom=None, diagonal=None, diagonal_direction=None, outline=true))]
    fn new(
        left: Option<PySide>,
        right: Option<PySide>,
        top: Option<PySide>,
        bottom: Option<PySide>,
        diagonal: Option<PySide>,
        diagonal_direction: Option<String>,
        outline: bool,
    ) -> Self {
        PyBorder {
            left,
            right,
            top,
            bottom,
            diagonal,
            diagonal_direction,
            outline,
        }
    }

    fn copy(&self) -> PyBorder {
        self.clone()
    }

    fn __str__(&self) -> String {
        format!(
            "<Border left={:?} right={:?} top={:?} bottom={:?}>",
            self.left.is_some(),
            self.right.is_some(),
            self.top.is_some(),
            self.bottom.is_some()
        )
    }

    fn __repr__(&self) -> String {
        self.__str__()
    }
}

/// Color (openpyxl-compatible).
#[pyclass(name = "Color")]
#[derive(Clone, Debug)]
pub struct PyColor {
    #[pyo3(get, set)]
    pub rgb: Option<String>,
    #[pyo3(get, set)]
    pub theme: Option<u32>,
    #[pyo3(get, set)]
    pub tint: f64,
    #[pyo3(get, set)]
    pub indexed: Option<u32>,
}

#[pymethods]
impl PyColor {
    #[new]
    #[pyo3(signature = (rgb=None, theme=None, tint=0.0, indexed=None))]
    fn new(
        rgb: Option<String>,
        theme: Option<u32>,
        tint: f64,
        indexed: Option<u32>,
    ) -> Self {
        PyColor {
            rgb,
            theme,
            tint,
            indexed,
        }
    }

    fn copy(&self) -> PyColor {
        self.clone()
    }

    fn __str__(&self) -> String {
        if let Some(ref rgb) = self.rgb {
            format!("<Color rgb={}>", rgb)
        } else if let Some(theme) = self.theme {
            format!("<Color theme={}>", theme)
        } else {
            "<Color>".to_string()
        }
    }

    fn __repr__(&self) -> String {
        self.__str__()
    }
}

impl Default for PyColor {
    fn default() -> Self {
        PyColor {
            rgb: None,
            theme: None,
            tint: 0.0,
            indexed: None,
        }
    }
}

/// Protection (openpyxl-compatible).
#[pyclass(name = "Protection")]
#[derive(Clone, Debug, Default)]
pub struct PyProtection {
    #[pyo3(get, set)]
    pub locked: bool,
    #[pyo3(get, set)]
    pub hidden: bool,
}

#[pymethods]
impl PyProtection {
    #[new]
    #[pyo3(signature = (locked=true, hidden=false))]
    fn new(locked: bool, hidden: bool) -> Self {
        PyProtection { locked, hidden }
    }

    fn copy(&self) -> PyProtection {
        self.clone()
    }

    fn __str__(&self) -> String {
        format!("<Protection locked={} hidden={}>", self.locked, self.hidden)
    }

    fn __repr__(&self) -> String {
        self.__str__()
    }
}
