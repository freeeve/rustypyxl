//! Python bindings for styling classes.

#![allow(non_snake_case)]

use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use rustypyxl_core::Color;

/// Accept either an rgb string or a Color object wherever openpyxl does.
pub(crate) fn coerce_color(value: Option<&Bound<'_, PyAny>>) -> PyResult<Option<Color>> {
    let Some(v) = value else { return Ok(None) };
    if v.is_none() {
        return Ok(None);
    }
    if let Ok(rgb) = v.extract::<String>() {
        return Ok(Some(Color::rgb(rgb)));
    }
    if let Ok(color) = v.extract::<PyColor>() {
        let color = Color {
            rgb: color.rgb,
            theme: color.theme,
            indexed: color.indexed,
            // 0.0 is the default, i.e. no tint at all
            tint: (color.tint != 0.0).then_some(color.tint),
        };
        return Ok((!color.is_empty()).then_some(color));
    }
    Err(PyTypeError::new_err(
        "expected an rgb string or a Color object",
    ))
}

/// Hand a color back to Python.
///
/// A color that is only an rgb value comes back as the plain hex string it has
/// always been, so existing code keeps working. One that carries a theme, a
/// palette index, or a tint comes back as a Color, because a string cannot
/// express those -- they used to be discarded on the way in.
pub(crate) fn color_to_python(color: Option<&Color>, py: Python<'_>) -> PyResult<PyObject> {
    let Some(color) = color else {
        return Ok(py.None());
    };

    if color.theme.is_none() && color.indexed.is_none() && color.tint.is_none() {
        if let Some(ref rgb) = color.rgb {
            return Ok(rgb.clone().into_pyobject(py)?.into_any().unbind());
        }
    }

    Ok(Py::new(
        py,
        PyColor {
            rgb: color.rgb.clone(),
            theme: color.theme,
            tint: color.tint.unwrap_or(0.0),
            indexed: color.indexed,
        },
    )?
    .into_any())
}

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
    /// Exposed through get_color/set_color, which accept and return either a
    /// hex string or a Color.
    pub color: Option<Color>,
    #[pyo3(get, set)]
    pub vertAlign: Option<String>,
}

#[pymethods]
impl PyFont {
    #[getter(color)]
    fn get_color(&self, py: Python<'_>) -> PyResult<PyObject> {
        color_to_python(self.color.as_ref(), py)
    }

    #[setter(color)]
    fn set_color(&mut self, value: Option<Bound<'_, PyAny>>) -> PyResult<()> {
        self.color = coerce_color(value.as_ref())?;
        Ok(())
    }

    #[new]
    // Mirrors openpyxl's Font keyword arguments
    #[allow(clippy::too_many_arguments)]
    #[pyo3(signature = (name=None, size=None, bold=false, italic=false, underline=None, strike=false, color=None, vertAlign=None))]
    fn new(
        name: Option<String>,
        size: Option<f64>,
        bold: bool,
        italic: bool,
        underline: Option<String>,
        strike: bool,
        color: Option<Bound<'_, PyAny>>,
        vertAlign: Option<String>,
    ) -> PyResult<Self> {
        Ok(PyFont {
            name,
            size,
            bold,
            italic,
            underline,
            strike,
            color: coerce_color(color.as_ref())?,
            vertAlign,
        })
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
    /// See PyFont::color -- a hex string, or a Color when it carries more.
    pub fgColor: Option<Color>,
    pub bgColor: Option<Color>,
    #[pyo3(get, set)]
    pub patternType: Option<String>,
}

#[pymethods]
impl PyPatternFill {
    #[getter(fgColor)]
    fn get_fg_color(&self, py: Python<'_>) -> PyResult<PyObject> {
        color_to_python(self.fgColor.as_ref(), py)
    }

    #[setter(fgColor)]
    fn set_fg_color(&mut self, value: Option<Bound<'_, PyAny>>) -> PyResult<()> {
        self.fgColor = coerce_color(value.as_ref())?;
        Ok(())
    }

    #[getter(bgColor)]
    fn get_bg_color(&self, py: Python<'_>) -> PyResult<PyObject> {
        color_to_python(self.bgColor.as_ref(), py)
    }

    #[setter(bgColor)]
    fn set_bg_color(&mut self, value: Option<Bound<'_, PyAny>>) -> PyResult<()> {
        self.bgColor = coerce_color(value.as_ref())?;
        Ok(())
    }

    #[new]
    // start_color/end_color are openpyxl's canonical aliases for fgColor/bgColor
    #[pyo3(signature = (fill_type=None, fgColor=None, bgColor=None, patternType=None, start_color=None, end_color=None))]
    fn new(
        fill_type: Option<String>,
        fgColor: Option<Bound<'_, PyAny>>,
        bgColor: Option<Bound<'_, PyAny>>,
        patternType: Option<String>,
        start_color: Option<Bound<'_, PyAny>>,
        end_color: Option<Bound<'_, PyAny>>,
    ) -> PyResult<Self> {
        let fg = coerce_color(fgColor.as_ref())?.or(coerce_color(start_color.as_ref())?);
        let bg = coerce_color(bgColor.as_ref())?.or(coerce_color(end_color.as_ref())?);
        Ok(PyPatternFill {
            fill_type: fill_type.or(patternType.clone()),
            fgColor: fg,
            bgColor: bg,
            patternType,
        })
    }

    /// openpyxl alias for fgColor.
    #[getter]
    fn start_color(&self, py: Python<'_>) -> PyResult<PyObject> {
        color_to_python(self.fgColor.as_ref(), py)
    }

    /// openpyxl alias for bgColor.
    #[getter]
    fn end_color(&self, py: Python<'_>) -> PyResult<PyObject> {
        color_to_python(self.bgColor.as_ref(), py)
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
    /// See PyFont::color -- a hex string, or a Color when it carries more.
    pub color: Option<Color>,
}

#[pymethods]
impl PySide {
    #[getter(color)]
    fn get_color(&self, py: Python<'_>) -> PyResult<PyObject> {
        color_to_python(self.color.as_ref(), py)
    }

    #[setter(color)]
    fn set_color(&mut self, value: Option<Bound<'_, PyAny>>) -> PyResult<()> {
        self.color = coerce_color(value.as_ref())?;
        Ok(())
    }

    #[new]
    #[pyo3(signature = (style=None, color=None))]
    fn new(style: Option<String>, color: Option<Bound<'_, PyAny>>) -> PyResult<Self> {
        Ok(PySide {
            style,
            color: coerce_color(color.as_ref())?,
        })
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
    fn new(rgb: Option<String>, theme: Option<u32>, tint: f64, indexed: Option<u32>) -> Self {
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

/// Gradient stop for gradient fills.
#[pyclass(name = "GradientStop")]
#[derive(Clone, Debug, Default)]
pub struct PyGradientStop {
    #[pyo3(get, set)]
    pub position: f64,
    #[pyo3(get, set)]
    pub color: Option<String>,
}

#[pymethods]
impl PyGradientStop {
    #[new]
    #[pyo3(signature = (position=0.0, color=None))]
    fn new(position: f64, color: Option<String>) -> Self {
        PyGradientStop { position, color }
    }

    fn copy(&self) -> PyGradientStop {
        self.clone()
    }

    fn __str__(&self) -> String {
        format!(
            "<GradientStop position={} color={:?}>",
            self.position, self.color
        )
    }

    fn __repr__(&self) -> String {
        self.__str__()
    }
}

/// Gradient fill (openpyxl-compatible).
#[pyclass(name = "GradientFill")]
#[derive(Clone, Debug, Default)]
pub struct PyGradientFill {
    #[pyo3(get, set)]
    pub fill_type: Option<String>,
    #[pyo3(get, set)]
    pub degree: Option<f64>,
    #[pyo3(get, set)]
    pub left: Option<f64>,
    #[pyo3(get, set)]
    pub right: Option<f64>,
    #[pyo3(get, set)]
    pub top: Option<f64>,
    #[pyo3(get, set)]
    pub bottom: Option<f64>,
    #[pyo3(get, set)]
    pub stop: Vec<PyGradientStop>,
}

#[pymethods]
impl PyGradientFill {
    #[new]
    #[pyo3(signature = (fill_type=None, degree=None, left=None, right=None, top=None, bottom=None, stop=None))]
    fn new(
        fill_type: Option<String>,
        degree: Option<f64>,
        left: Option<f64>,
        right: Option<f64>,
        top: Option<f64>,
        bottom: Option<f64>,
        stop: Option<Vec<PyGradientStop>>,
    ) -> Self {
        PyGradientFill {
            fill_type: fill_type.or(Some("linear".to_string())),
            degree: degree.or(Some(0.0)),
            left,
            right,
            top,
            bottom,
            stop: stop.unwrap_or_default(),
        }
    }

    fn copy(&self) -> PyGradientFill {
        self.clone()
    }

    fn __str__(&self) -> String {
        format!(
            "<GradientFill type={:?} degree={:?} stops={}>",
            self.fill_type,
            self.degree,
            self.stop.len()
        )
    }

    fn __repr__(&self) -> String {
        self.__str__()
    }
}
