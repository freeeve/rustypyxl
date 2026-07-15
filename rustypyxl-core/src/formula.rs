//! A small formula evaluation engine: tokenizer, precedence parser, and
//! evaluator for the common subset of Excel formulas.
//!
//! **Scope (v1).** Arithmetic (`+ - * / ^`), string concat (`&`), comparisons
//! (`= <> < > <= >=`), unary minus and trailing `%`, numbers, quoted strings,
//! booleans, cell references (`A1`, `$A$1`), same-sheet ranges (`A1:B10`),
//! sheet-qualified references (`Sheet!A1`, `'My Sheet'!A1`), and a starter set
//! of functions: SUM, AVERAGE, COUNT, COUNTA, MIN, MAX, PRODUCT, ROUND, ABS,
//! SQRT, INT, MOD, POWER, SUMIF, COUNTIF, IF, AND, OR, NOT, CONCATENATE, LEN,
//! LEFT, RIGHT, MID, UPPER, LOWER, TRIM, TRUE, FALSE.
//!
//! Anything outside this subset resolves to an Excel-style error value
//! (`#NAME?` for an unknown function, `#VALUE!` for a type error, `#DIV/0!`,
//! `#REF!`); it never panics. Cell references are resolved through a
//! [`CellResolver`], so the same engine works over a worksheet or any other
//! backing store.

/// A value produced by evaluating a formula (or read from a referenced cell).
#[derive(Clone, Debug, PartialEq)]
pub enum FormulaValue {
    /// A number.
    Number(f64),
    /// Text.
    Text(String),
    /// A boolean.
    Bool(bool),
    /// A blank cell.
    Empty,
    /// An Excel error value, e.g. "#DIV/0!", "#VALUE!", "#NAME?", "#REF!".
    Error(String),
}

impl FormulaValue {
    /// Coerce to a number for arithmetic (Excel rules: blanks and FALSE are 0,
    /// TRUE is 1, numeric text parses, other text is a #VALUE! error).
    fn to_number(&self) -> Result<f64, FormulaValue> {
        match self {
            FormulaValue::Number(n) => Ok(*n),
            FormulaValue::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
            FormulaValue::Empty => Ok(0.0),
            FormulaValue::Text(t) => t
                .trim()
                .parse::<f64>()
                .map_err(|_| FormulaValue::Error("#VALUE!".to_string())),
            FormulaValue::Error(_) => Err(self.clone()),
        }
    }

    /// Coerce to text (numbers use a plain representation; blanks are empty).
    fn to_text(&self) -> Result<String, FormulaValue> {
        match self {
            FormulaValue::Text(t) => Ok(t.clone()),
            FormulaValue::Number(n) => Ok(format_number_plain(*n)),
            FormulaValue::Bool(b) => Ok(if *b { "TRUE" } else { "FALSE" }.to_string()),
            FormulaValue::Empty => Ok(String::new()),
            FormulaValue::Error(_) => Err(self.clone()),
        }
    }

    /// Coerce to a boolean for logical context.
    fn to_bool(&self) -> Result<bool, FormulaValue> {
        match self {
            FormulaValue::Bool(b) => Ok(*b),
            FormulaValue::Number(n) => Ok(*n != 0.0),
            FormulaValue::Empty => Ok(false),
            FormulaValue::Text(t) => match t.to_ascii_uppercase().as_str() {
                "TRUE" => Ok(true),
                "FALSE" => Ok(false),
                _ => Err(FormulaValue::Error("#VALUE!".to_string())),
            },
            FormulaValue::Error(_) => Err(self.clone()),
        }
    }

    /// Whether this is an error value.
    pub fn is_error(&self) -> bool {
        matches!(self, FormulaValue::Error(_))
    }
}

/// Format a number without a trailing ".0" for integers, matching how Excel
/// coerces a number to text.
fn format_number_plain(n: f64) -> String {
    if n == n.trunc() && n.abs() < 1e15 {
        format!("{}", n as i64)
    } else {
        format!("{}", n)
    }
}

/// Resolves a cell reference to a value. Sheet is `None` for an unqualified
/// reference (the current sheet).
pub trait CellResolver {
    /// Resolve the 1-based (row, col) cell on the given sheet.
    fn resolve(&mut self, sheet: Option<&str>, row: u32, col: u32) -> FormulaValue;
}

// ---------- tokenizer ----------

#[derive(Clone, Debug, PartialEq)]
enum Token {
    Num(f64),
    Str(String),
    Bool(bool),
    /// A cell or range reference, sheet-qualified if it contains '!'.
    Ref(String),
    /// A function name (was followed by '(').
    Func(String),
    Op(String),
    LParen,
    RParen,
    Comma,
    Percent,
}

fn tokenize(input: &str) -> Result<Vec<Token>, FormulaValue> {
    let chars: Vec<char> = input.chars().collect();
    let mut i = 0;
    let mut tokens = Vec::new();

    while i < chars.len() {
        let c = chars[i];
        match c {
            ' ' | '\t' | '\n' | '\r' => i += 1,
            '(' => {
                tokens.push(Token::LParen);
                i += 1;
            }
            ')' => {
                tokens.push(Token::RParen);
                i += 1;
            }
            ',' => {
                tokens.push(Token::Comma);
                i += 1;
            }
            '%' => {
                tokens.push(Token::Percent);
                i += 1;
            }
            '+' | '-' | '*' | '/' | '^' | '&' | '=' => {
                tokens.push(Token::Op(c.to_string()));
                i += 1;
            }
            '<' => {
                if chars.get(i + 1) == Some(&'=') {
                    tokens.push(Token::Op("<=".to_string()));
                    i += 2;
                } else if chars.get(i + 1) == Some(&'>') {
                    tokens.push(Token::Op("<>".to_string()));
                    i += 2;
                } else {
                    tokens.push(Token::Op("<".to_string()));
                    i += 1;
                }
            }
            '>' => {
                if chars.get(i + 1) == Some(&'=') {
                    tokens.push(Token::Op(">=".to_string()));
                    i += 2;
                } else {
                    tokens.push(Token::Op(">".to_string()));
                    i += 1;
                }
            }
            '"' => {
                let mut s = String::new();
                i += 1;
                while i < chars.len() {
                    if chars[i] == '"' {
                        if chars.get(i + 1) == Some(&'"') {
                            s.push('"');
                            i += 2;
                        } else {
                            i += 1;
                            break;
                        }
                    } else {
                        s.push(chars[i]);
                        i += 1;
                    }
                }
                tokens.push(Token::Str(s));
            }
            '\'' => {
                // quoted sheet name: 'My Sheet'!A1[:B2]
                let mut name = String::from("'");
                i += 1;
                while i < chars.len() && chars[i] != '\'' {
                    name.push(chars[i]);
                    i += 1;
                }
                name.push('\'');
                i += 1; // closing quote
                let mut reference = name;
                while i < chars.len() && is_ref_char(chars[i]) {
                    reference.push(chars[i]);
                    i += 1;
                }
                tokens.push(Token::Ref(reference));
            }
            c if c.is_ascii_digit()
                || (c == '.' && chars.get(i + 1).is_some_and(|d| d.is_ascii_digit())) =>
            {
                let start = i;
                while i < chars.len()
                    && (chars[i].is_ascii_digit()
                        || chars[i] == '.'
                        || chars[i] == 'e'
                        || chars[i] == 'E'
                        || ((chars[i] == '+' || chars[i] == '-')
                            && matches!(chars.get(i - 1), Some('e') | Some('E'))))
                {
                    i += 1;
                }
                let num: String = chars[start..i].iter().collect();
                let n = num
                    .parse::<f64>()
                    .map_err(|_| FormulaValue::Error("#VALUE!".to_string()))?;
                tokens.push(Token::Num(n));
            }
            c if c.is_ascii_alphabetic() || c == '$' || c == '_' => {
                let start = i;
                while i < chars.len() && (is_ref_char(chars[i]) || chars[i] == '$') {
                    i += 1;
                }
                // sheet-qualified reference: identifier '!' cellref
                if i < chars.len() && chars[i] == '!' {
                    i += 1;
                    let mut reference: String = chars[start..i].iter().collect();
                    while i < chars.len() && is_ref_char(chars[i]) {
                        reference.push(chars[i]);
                        i += 1;
                    }
                    tokens.push(Token::Ref(reference));
                    continue;
                }
                let word: String = chars[start..i].iter().collect();
                // skip spaces to detect a following '('
                let mut j = i;
                while j < chars.len() && chars[j] == ' ' {
                    j += 1;
                }
                if chars.get(j) == Some(&'(') {
                    tokens.push(Token::Func(word.to_ascii_uppercase()));
                } else if word.eq_ignore_ascii_case("TRUE") {
                    tokens.push(Token::Bool(true));
                } else if word.eq_ignore_ascii_case("FALSE") {
                    tokens.push(Token::Bool(false));
                } else if looks_like_ref(&word) {
                    tokens.push(Token::Ref(word));
                } else {
                    // named ranges and defined names are unsupported in v1
                    return Err(FormulaValue::Error("#NAME?".to_string()));
                }
            }
            _ => return Err(FormulaValue::Error("#VALUE!".to_string())),
        }
    }
    Ok(tokens)
}

fn is_ref_char(c: char) -> bool {
    c.is_ascii_alphanumeric() || c == '$' || c == ':' || c == '.'
}

/// Whether a bare word (no sheet prefix) is a plain A1-style cell reference or a
/// range of two such references.
fn looks_like_ref(word: &str) -> bool {
    match word.split_once(':') {
        Some((a, b)) => is_cell_word(a) && is_cell_word(b),
        None => is_cell_word(word),
    }
}

/// Whether a word is a single A1-style cell reference (letters then digits,
/// with optional `$` anchors).
fn is_cell_word(word: &str) -> bool {
    let w = word.replace('$', "");
    if w.is_empty() {
        return false;
    }
    let mut seen_letter = false;
    let mut seen_digit = false;
    for c in w.chars() {
        if c.is_ascii_alphabetic() && !seen_digit {
            seen_letter = true;
        } else if c.is_ascii_digit() {
            seen_digit = true;
        } else {
            return false;
        }
    }
    seen_letter && seen_digit
}

// ---------- parser (recursive descent with precedence) ----------

#[derive(Clone, Debug)]
enum Expr {
    Num(f64),
    Text(String),
    Bool(bool),
    Cell {
        sheet: Option<String>,
        row: u32,
        col: u32,
    },
    Range {
        sheet: Option<String>,
        r1: u32,
        c1: u32,
        r2: u32,
        c2: u32,
    },
    Unary(String, Box<Expr>),
    Binary(String, Box<Expr>, Box<Expr>),
    Func(String, Vec<Expr>),
}

struct Parser {
    tokens: Vec<Token>,
    pos: usize,
}

impl Parser {
    fn peek(&self) -> Option<&Token> {
        self.tokens.get(self.pos)
    }
    fn next(&mut self) -> Option<Token> {
        let t = self.tokens.get(self.pos).cloned();
        self.pos += 1;
        t
    }

    /// comparison is lowest precedence
    fn parse_expr(&mut self) -> Result<Expr, FormulaValue> {
        let mut left = self.parse_concat()?;
        while let Some(Token::Op(op)) = self.peek() {
            if matches!(op.as_str(), "=" | "<>" | "<" | ">" | "<=" | ">=") {
                let op = op.clone();
                self.pos += 1;
                let right = self.parse_concat()?;
                left = Expr::Binary(op, Box::new(left), Box::new(right));
            } else {
                break;
            }
        }
        Ok(left)
    }

    fn parse_concat(&mut self) -> Result<Expr, FormulaValue> {
        let mut left = self.parse_add()?;
        while let Some(Token::Op(op)) = self.peek() {
            if op == "&" {
                self.pos += 1;
                let right = self.parse_add()?;
                left = Expr::Binary("&".to_string(), Box::new(left), Box::new(right));
            } else {
                break;
            }
        }
        Ok(left)
    }

    fn parse_add(&mut self) -> Result<Expr, FormulaValue> {
        let mut left = self.parse_mul()?;
        while let Some(Token::Op(op)) = self.peek() {
            if op == "+" || op == "-" {
                let op = op.clone();
                self.pos += 1;
                let right = self.parse_mul()?;
                left = Expr::Binary(op, Box::new(left), Box::new(right));
            } else {
                break;
            }
        }
        Ok(left)
    }

    fn parse_mul(&mut self) -> Result<Expr, FormulaValue> {
        let mut left = self.parse_pow()?;
        while let Some(Token::Op(op)) = self.peek() {
            if op == "*" || op == "/" {
                let op = op.clone();
                self.pos += 1;
                let right = self.parse_pow()?;
                left = Expr::Binary(op, Box::new(left), Box::new(right));
            } else {
                break;
            }
        }
        Ok(left)
    }

    fn parse_pow(&mut self) -> Result<Expr, FormulaValue> {
        let left = self.parse_unary()?;
        if let Some(Token::Op(op)) = self.peek() {
            if op == "^" {
                self.pos += 1;
                let right = self.parse_pow()?; // right-associative
                return Ok(Expr::Binary(
                    "^".to_string(),
                    Box::new(left),
                    Box::new(right),
                ));
            }
        }
        Ok(left)
    }

    fn parse_unary(&mut self) -> Result<Expr, FormulaValue> {
        if let Some(Token::Op(op)) = self.peek() {
            if op == "-" || op == "+" {
                let op = op.clone();
                self.pos += 1;
                let operand = self.parse_unary()?;
                return Ok(Expr::Unary(op, Box::new(operand)));
            }
        }
        self.parse_postfix()
    }

    fn parse_postfix(&mut self) -> Result<Expr, FormulaValue> {
        let mut expr = self.parse_atom()?;
        while matches!(self.peek(), Some(Token::Percent)) {
            self.pos += 1;
            expr = Expr::Unary("%".to_string(), Box::new(expr));
        }
        Ok(expr)
    }

    fn parse_atom(&mut self) -> Result<Expr, FormulaValue> {
        match self.next() {
            Some(Token::Num(n)) => Ok(Expr::Num(n)),
            Some(Token::Str(s)) => Ok(Expr::Text(s)),
            Some(Token::Bool(b)) => Ok(Expr::Bool(b)),
            Some(Token::Ref(r)) => parse_reference(&r),
            Some(Token::LParen) => {
                let e = self.parse_expr()?;
                match self.next() {
                    Some(Token::RParen) => Ok(e),
                    _ => Err(FormulaValue::Error("#VALUE!".to_string())),
                }
            }
            Some(Token::Func(name)) => {
                // consume '('
                if !matches!(self.next(), Some(Token::LParen)) {
                    return Err(FormulaValue::Error("#VALUE!".to_string()));
                }
                let mut args = Vec::new();
                if !matches!(self.peek(), Some(Token::RParen)) {
                    loop {
                        args.push(self.parse_expr()?);
                        match self.peek() {
                            Some(Token::Comma) => {
                                self.pos += 1;
                            }
                            _ => break,
                        }
                    }
                }
                match self.next() {
                    Some(Token::RParen) => Ok(Expr::Func(name, args)),
                    _ => Err(FormulaValue::Error("#VALUE!".to_string())),
                }
            }
            _ => Err(FormulaValue::Error("#VALUE!".to_string())),
        }
    }
}

/// Parse a reference token (`A1`, `$A$1`, `A1:B2`, `Sheet!A1`, `'S'!A1:B2`).
fn parse_reference(reference: &str) -> Result<Expr, FormulaValue> {
    let (sheet, cells) = match reference.rfind('!') {
        Some(idx) => {
            let mut name = reference[..idx].to_string();
            if name.starts_with('\'') && name.ends_with('\'') && name.len() >= 2 {
                name = name[1..name.len() - 1].replace("''", "'");
            }
            (Some(name), &reference[idx + 1..])
        }
        None => (None, reference),
    };

    if let Some((a, b)) = cells.split_once(':') {
        let (r1, c1) = parse_a1(a).ok_or_else(|| FormulaValue::Error("#REF!".to_string()))?;
        let (r2, c2) = parse_a1(b).ok_or_else(|| FormulaValue::Error("#REF!".to_string()))?;
        Ok(Expr::Range {
            sheet,
            r1: r1.min(r2),
            c1: c1.min(c2),
            r2: r1.max(r2),
            c2: c1.max(c2),
        })
    } else {
        let (row, col) = parse_a1(cells).ok_or_else(|| FormulaValue::Error("#REF!".to_string()))?;
        Ok(Expr::Cell { sheet, row, col })
    }
}

/// Parse an `A1` / `$A$1` cell into 1-based (row, col).
fn parse_a1(s: &str) -> Option<(u32, u32)> {
    let s = s.replace('$', "");
    crate::utils::parse_coordinate(&s).ok()
}

// ---------- evaluator ----------

/// Parse and evaluate a formula (with or without a leading `=`) against a
/// resolver, returning the resulting value. Never panics; malformed input and
/// type errors resolve to Excel-style error values.
pub fn evaluate(formula: &str, resolver: &mut dyn CellResolver) -> FormulaValue {
    let body = formula.strip_prefix('=').unwrap_or(formula);
    let tokens = match tokenize(body) {
        Ok(t) => t,
        Err(e) => return e,
    };
    if tokens.is_empty() {
        return FormulaValue::Empty;
    }
    let mut parser = Parser { tokens, pos: 0 };
    let expr = match parser.parse_expr() {
        Ok(e) => e,
        Err(e) => return e,
    };
    if parser.pos != parser.tokens.len() {
        return FormulaValue::Error("#VALUE!".to_string());
    }
    eval_expr(&expr, resolver)
}

/// Evaluate an expression to a single scalar value.
fn eval_expr(expr: &Expr, resolver: &mut dyn CellResolver) -> FormulaValue {
    match expr {
        Expr::Num(n) => FormulaValue::Number(*n),
        Expr::Text(t) => FormulaValue::Text(t.clone()),
        Expr::Bool(b) => FormulaValue::Bool(*b),
        Expr::Cell { sheet, row, col } => resolver.resolve(sheet.as_deref(), *row, *col),
        Expr::Range { .. } => FormulaValue::Error("#VALUE!".to_string()), // range in scalar context
        Expr::Unary(op, e) => {
            let v = eval_expr(e, resolver);
            match op.as_str() {
                "-" => match v.to_number() {
                    Ok(n) => FormulaValue::Number(-n),
                    Err(e) => e,
                },
                "+" => match v.to_number() {
                    Ok(n) => FormulaValue::Number(n),
                    Err(e) => e,
                },
                "%" => match v.to_number() {
                    Ok(n) => FormulaValue::Number(n / 100.0),
                    Err(e) => e,
                },
                _ => FormulaValue::Error("#VALUE!".to_string()),
            }
        }
        Expr::Binary(op, l, r) => eval_binary(op, l, r, resolver),
        Expr::Func(name, args) => eval_function(name, args, resolver),
    }
}

fn eval_binary(op: &str, l: &Expr, r: &Expr, resolver: &mut dyn CellResolver) -> FormulaValue {
    let lv = eval_expr(l, resolver);
    let rv = eval_expr(r, resolver);
    if lv.is_error() {
        return lv;
    }
    if rv.is_error() {
        return rv;
    }

    if op == "&" {
        let a = match lv.to_text() {
            Ok(t) => t,
            Err(e) => return e,
        };
        let b = match rv.to_text() {
            Ok(t) => t,
            Err(e) => return e,
        };
        return FormulaValue::Text(a + &b);
    }

    if matches!(op, "=" | "<>" | "<" | ">" | "<=" | ">=") {
        return compare(op, &lv, &rv);
    }

    let a = match lv.to_number() {
        Ok(n) => n,
        Err(e) => return e,
    };
    let b = match rv.to_number() {
        Ok(n) => n,
        Err(e) => return e,
    };
    match op {
        "+" => FormulaValue::Number(a + b),
        "-" => FormulaValue::Number(a - b),
        "*" => FormulaValue::Number(a * b),
        "/" => {
            if b == 0.0 {
                FormulaValue::Error("#DIV/0!".to_string())
            } else {
                FormulaValue::Number(a / b)
            }
        }
        "^" => FormulaValue::Number(a.powf(b)),
        _ => FormulaValue::Error("#VALUE!".to_string()),
    }
}

/// Compare two values with Excel semantics: numbers numerically, text
/// case-insensitively, with numbers sorting before text.
fn compare(op: &str, l: &FormulaValue, r: &FormulaValue) -> FormulaValue {
    use std::cmp::Ordering;
    let ord = match (l, r) {
        (FormulaValue::Number(a), FormulaValue::Number(b)) => {
            a.partial_cmp(b).unwrap_or(Ordering::Equal)
        }
        (FormulaValue::Empty, FormulaValue::Number(b)) => {
            0.0f64.partial_cmp(b).unwrap_or(Ordering::Equal)
        }
        (FormulaValue::Number(a), FormulaValue::Empty) => {
            a.partial_cmp(&0.0).unwrap_or(Ordering::Equal)
        }
        (FormulaValue::Bool(a), FormulaValue::Bool(b)) => a.cmp(b),
        _ => {
            // fall back to case-insensitive text comparison
            let a = l.to_text().unwrap_or_default().to_ascii_uppercase();
            let b = r.to_text().unwrap_or_default().to_ascii_uppercase();
            a.cmp(&b)
        }
    };
    let result = match op {
        "=" => ord == Ordering::Equal,
        "<>" => ord != Ordering::Equal,
        "<" => ord == Ordering::Less,
        ">" => ord == Ordering::Greater,
        "<=" => ord != Ordering::Greater,
        ">=" => ord != Ordering::Less,
        _ => return FormulaValue::Error("#VALUE!".to_string()),
    };
    FormulaValue::Bool(result)
}

/// Expand a function argument into a flat list of values (a range yields every
/// cell; a scalar yields one). Errors propagate as a single-element list.
fn eval_arg_values(expr: &Expr, resolver: &mut dyn CellResolver) -> Vec<FormulaValue> {
    match expr {
        Expr::Range {
            sheet,
            r1,
            c1,
            r2,
            c2,
        } => {
            let mut out = Vec::new();
            for row in *r1..=*r2 {
                for col in *c1..=*c2 {
                    out.push(resolver.resolve(sheet.as_deref(), row, col));
                }
            }
            out
        }
        _ => vec![eval_expr(expr, resolver)],
    }
}

/// Collect the numbers from a set of function arguments, ignoring text and
/// blanks (Excel's aggregate behavior). Propagates the first error.
fn collect_numbers(
    args: &[Expr],
    resolver: &mut dyn CellResolver,
) -> Result<Vec<f64>, FormulaValue> {
    let mut nums = Vec::new();
    for arg in args {
        for v in eval_arg_values(arg, resolver) {
            match v {
                FormulaValue::Number(n) => nums.push(n),
                FormulaValue::Bool(b) => nums.push(if b { 1.0 } else { 0.0 }),
                FormulaValue::Error(_) => return Err(v),
                // text and blanks are ignored by SUM/AVERAGE/etc.
                _ => {}
            }
        }
    }
    Ok(nums)
}

fn eval_function(name: &str, args: &[Expr], resolver: &mut dyn CellResolver) -> FormulaValue {
    match name {
        "SUM" => match collect_numbers(args, resolver) {
            Ok(nums) => FormulaValue::Number(nums.iter().sum()),
            Err(e) => e,
        },
        "PRODUCT" => match collect_numbers(args, resolver) {
            Ok(nums) => FormulaValue::Number(nums.iter().product()),
            Err(e) => e,
        },
        "AVERAGE" => match collect_numbers(args, resolver) {
            Ok(nums) if nums.is_empty() => FormulaValue::Error("#DIV/0!".to_string()),
            Ok(nums) => FormulaValue::Number(nums.iter().sum::<f64>() / nums.len() as f64),
            Err(e) => e,
        },
        "MIN" => match collect_numbers(args, resolver) {
            Ok(nums) if nums.is_empty() => FormulaValue::Number(0.0),
            Ok(nums) => FormulaValue::Number(nums.iter().cloned().fold(f64::INFINITY, f64::min)),
            Err(e) => e,
        },
        "MAX" => match collect_numbers(args, resolver) {
            Ok(nums) if nums.is_empty() => FormulaValue::Number(0.0),
            Ok(nums) => {
                FormulaValue::Number(nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max))
            }
            Err(e) => e,
        },
        "COUNT" => {
            let mut count = 0i64;
            for arg in args {
                for v in eval_arg_values(arg, resolver) {
                    if matches!(v, FormulaValue::Number(_)) {
                        count += 1;
                    }
                }
            }
            FormulaValue::Number(count as f64)
        }
        "COUNTA" => {
            let mut count = 0i64;
            for arg in args {
                for v in eval_arg_values(arg, resolver) {
                    if !matches!(v, FormulaValue::Empty) {
                        count += 1;
                    }
                }
            }
            FormulaValue::Number(count as f64)
        }
        "IF" => {
            if args.len() < 2 || args.len() > 3 {
                return FormulaValue::Error("#VALUE!".to_string());
            }
            match eval_expr(&args[0], resolver).to_bool() {
                Ok(true) => eval_expr(&args[1], resolver),
                Ok(false) => {
                    if args.len() == 3 {
                        eval_expr(&args[2], resolver)
                    } else {
                        FormulaValue::Bool(false)
                    }
                }
                Err(e) => e,
            }
        }
        "AND" | "OR" => {
            let mut result = name == "AND";
            let mut any = false;
            for arg in args {
                for v in eval_arg_values(arg, resolver) {
                    if matches!(v, FormulaValue::Empty | FormulaValue::Text(_)) {
                        continue;
                    }
                    any = true;
                    match v.to_bool() {
                        Ok(b) => {
                            if name == "AND" {
                                result = result && b;
                            } else {
                                result = result || b;
                            }
                        }
                        Err(e) => return e,
                    }
                }
            }
            if !any {
                return FormulaValue::Error("#VALUE!".to_string());
            }
            FormulaValue::Bool(result)
        }
        "NOT" => match single_arg(args, resolver) {
            Ok(v) => match v.to_bool() {
                Ok(b) => FormulaValue::Bool(!b),
                Err(e) => e,
            },
            Err(e) => e,
        },
        "ABS" => unary_num(args, resolver, f64::abs),
        "SQRT" => unary_num(args, resolver, |n| n.sqrt()),
        "INT" => unary_num(args, resolver, f64::floor),
        "ROUND" => {
            if args.len() != 2 {
                return FormulaValue::Error("#VALUE!".to_string());
            }
            let n = match num_of(&args[0], resolver) {
                Ok(n) => n,
                Err(e) => return e,
            };
            let digits = match num_of(&args[1], resolver) {
                Ok(d) => d,
                Err(e) => return e,
            };
            let factor = 10f64.powi(digits as i32);
            FormulaValue::Number((n * factor).round() / factor)
        }
        "MOD" => two_num(args, resolver, |a, b| {
            if b == 0.0 {
                FormulaValue::Error("#DIV/0!".to_string())
            } else {
                FormulaValue::Number(a - b * (a / b).floor())
            }
        }),
        "POWER" => two_num(args, resolver, |a, b| FormulaValue::Number(a.powf(b))),
        "SUMIF" => sumif(args, resolver),
        "COUNTIF" => countif(args, resolver),
        "CONCATENATE" => {
            let mut out = String::new();
            for arg in args {
                for v in eval_arg_values(arg, resolver) {
                    match v.to_text() {
                        Ok(t) => out.push_str(&t),
                        Err(e) => return e,
                    }
                }
            }
            FormulaValue::Text(out)
        }
        "LEN" => match single_arg(args, resolver) {
            Ok(v) => match v.to_text() {
                Ok(t) => FormulaValue::Number(t.chars().count() as f64),
                Err(e) => e,
            },
            Err(e) => e,
        },
        "UPPER" | "LOWER" | "TRIM" => match single_arg(args, resolver) {
            Ok(v) => match v.to_text() {
                Ok(t) => FormulaValue::Text(match name {
                    "UPPER" => t.to_uppercase(),
                    "LOWER" => t.to_lowercase(),
                    _ => t.trim().to_string(),
                }),
                Err(e) => e,
            },
            Err(e) => e,
        },
        "LEFT" | "RIGHT" => text_take(name, args, resolver),
        "MID" => mid(args, resolver),
        "TRUE" => FormulaValue::Bool(true),
        "FALSE" => FormulaValue::Bool(false),
        _ => FormulaValue::Error("#NAME?".to_string()),
    }
}

fn single_arg(
    args: &[Expr],
    resolver: &mut dyn CellResolver,
) -> Result<FormulaValue, FormulaValue> {
    if args.len() != 1 {
        return Err(FormulaValue::Error("#VALUE!".to_string()));
    }
    Ok(eval_expr(&args[0], resolver))
}

fn num_of(expr: &Expr, resolver: &mut dyn CellResolver) -> Result<f64, FormulaValue> {
    eval_expr(expr, resolver).to_number()
}

fn unary_num(args: &[Expr], resolver: &mut dyn CellResolver, f: fn(f64) -> f64) -> FormulaValue {
    match single_arg(args, resolver) {
        Ok(v) => match v.to_number() {
            Ok(n) => FormulaValue::Number(f(n)),
            Err(e) => e,
        },
        Err(e) => e,
    }
}

fn two_num(
    args: &[Expr],
    resolver: &mut dyn CellResolver,
    f: fn(f64, f64) -> FormulaValue,
) -> FormulaValue {
    if args.len() != 2 {
        return FormulaValue::Error("#VALUE!".to_string());
    }
    let a = match num_of(&args[0], resolver) {
        Ok(n) => n,
        Err(e) => return e,
    };
    let b = match num_of(&args[1], resolver) {
        Ok(n) => n,
        Err(e) => return e,
    };
    f(a, b)
}

fn text_take(name: &str, args: &[Expr], resolver: &mut dyn CellResolver) -> FormulaValue {
    if args.is_empty() || args.len() > 2 {
        return FormulaValue::Error("#VALUE!".to_string());
    }
    let text = match eval_expr(&args[0], resolver).to_text() {
        Ok(t) => t,
        Err(e) => return e,
    };
    let n = if args.len() == 2 {
        match num_of(&args[1], resolver) {
            Ok(n) => n.max(0.0) as usize,
            Err(e) => return e,
        }
    } else {
        1
    };
    let chars: Vec<char> = text.chars().collect();
    let taken: String = if name == "LEFT" {
        chars.iter().take(n).collect()
    } else {
        let start = chars.len().saturating_sub(n);
        chars[start..].iter().collect()
    };
    FormulaValue::Text(taken)
}

fn mid(args: &[Expr], resolver: &mut dyn CellResolver) -> FormulaValue {
    if args.len() != 3 {
        return FormulaValue::Error("#VALUE!".to_string());
    }
    let text = match eval_expr(&args[0], resolver).to_text() {
        Ok(t) => t,
        Err(e) => return e,
    };
    let start = match num_of(&args[1], resolver) {
        Ok(n) => n,
        Err(e) => return e,
    };
    let len = match num_of(&args[2], resolver) {
        Ok(n) => n,
        Err(e) => return e,
    };
    if start < 1.0 || len < 0.0 {
        return FormulaValue::Error("#VALUE!".to_string());
    }
    let chars: Vec<char> = text.chars().collect();
    let start_idx = (start as usize) - 1;
    let taken: String = chars.iter().skip(start_idx).take(len as usize).collect();
    FormulaValue::Text(taken)
}

/// SUMIF(range, criteria, [sum_range]) and COUNTIF share criteria matching.
fn sumif(args: &[Expr], resolver: &mut dyn CellResolver) -> FormulaValue {
    if args.len() != 2 && args.len() != 3 {
        return FormulaValue::Error("#VALUE!".to_string());
    }
    let range = eval_arg_values(&args[0], resolver);
    let criteria = match eval_expr(&args[1], resolver).to_text() {
        Ok(t) => t,
        Err(e) => return e,
    };
    let sum_range = if args.len() == 3 {
        eval_arg_values(&args[2], resolver)
    } else {
        range.clone()
    };
    let mut total = 0.0;
    for (i, v) in range.iter().enumerate() {
        if criteria_matches(v, &criteria) {
            if let Some(FormulaValue::Number(n)) = sum_range.get(i) {
                total += n;
            }
        }
    }
    FormulaValue::Number(total)
}

fn countif(args: &[Expr], resolver: &mut dyn CellResolver) -> FormulaValue {
    if args.len() != 2 {
        return FormulaValue::Error("#VALUE!".to_string());
    }
    let range = eval_arg_values(&args[0], resolver);
    let criteria = match eval_expr(&args[1], resolver).to_text() {
        Ok(t) => t,
        Err(e) => return e,
    };
    let count = range
        .iter()
        .filter(|v| criteria_matches(v, &criteria))
        .count();
    FormulaValue::Number(count as f64)
}

/// Match a value against a SUMIF/COUNTIF criterion: a comparison operator
/// prefix (>, <, >=, <=, <>, =) followed by a number, or an exact
/// number/text match.
fn criteria_matches(value: &FormulaValue, criteria: &str) -> bool {
    let criteria = criteria.trim();
    let (op, rest) = if let Some(r) = criteria.strip_prefix(">=") {
        (">=", r)
    } else if let Some(r) = criteria.strip_prefix("<=") {
        ("<=", r)
    } else if let Some(r) = criteria.strip_prefix("<>") {
        ("<>", r)
    } else if let Some(r) = criteria.strip_prefix('>') {
        (">", r)
    } else if let Some(r) = criteria.strip_prefix('<') {
        ("<", r)
    } else if let Some(r) = criteria.strip_prefix('=') {
        ("=", r)
    } else {
        ("=", criteria)
    };

    if let Ok(target) = rest.trim().parse::<f64>() {
        if let FormulaValue::Number(n) = value {
            return match op {
                ">" => *n > target,
                "<" => *n < target,
                ">=" => *n >= target,
                "<=" => *n <= target,
                "<>" => *n != target,
                _ => *n == target,
            };
        }
        return op == "<>";
    }

    // text criterion: exact, case-insensitive
    let text = value.to_text().unwrap_or_default();
    let eq = text.eq_ignore_ascii_case(rest.trim());
    if op == "<>" {
        !eq
    } else {
        eq
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;

    /// A resolver backed by a fixed map of (sheet, row, col) -> value.
    struct MapResolver {
        cells: HashMap<(Option<String>, u32, u32), FormulaValue>,
    }

    impl MapResolver {
        fn new() -> Self {
            MapResolver {
                cells: HashMap::new(),
            }
        }
        fn set(&mut self, cell: &str, v: FormulaValue) -> &mut Self {
            let (row, col) = parse_a1(cell).unwrap();
            self.cells.insert((None, row, col), v);
            self
        }
    }

    impl CellResolver for MapResolver {
        fn resolve(&mut self, sheet: Option<&str>, row: u32, col: u32) -> FormulaValue {
            self.cells
                .get(&(sheet.map(|s| s.to_string()), row, col))
                .cloned()
                .unwrap_or(FormulaValue::Empty)
        }
    }

    fn ev(formula: &str, r: &mut MapResolver) -> FormulaValue {
        evaluate(formula, r)
    }

    #[test]
    fn arithmetic_and_precedence() {
        let mut r = MapResolver::new();
        assert_eq!(ev("=1+2*3", &mut r), FormulaValue::Number(7.0));
        assert_eq!(ev("=(1+2)*3", &mut r), FormulaValue::Number(9.0));
        assert_eq!(ev("=2^3^2", &mut r), FormulaValue::Number(512.0)); // right assoc
        assert_eq!(ev("=-2^2", &mut r), FormulaValue::Number(4.0)); // unary binds looser than ^? Excel: -2^2 = 4
        assert_eq!(ev("=10/4", &mut r), FormulaValue::Number(2.5));
        assert_eq!(ev("=50%", &mut r), FormulaValue::Number(0.5));
    }

    #[test]
    fn division_by_zero() {
        let mut r = MapResolver::new();
        assert_eq!(ev("=1/0", &mut r), FormulaValue::Error("#DIV/0!".into()));
    }

    #[test]
    fn references_and_ranges() {
        let mut r = MapResolver::new();
        r.set("A1", FormulaValue::Number(10.0))
            .set("A2", FormulaValue::Number(20.0))
            .set("A3", FormulaValue::Number(30.0));
        assert_eq!(ev("=A1+A2", &mut r), FormulaValue::Number(30.0));
        assert_eq!(ev("=SUM(A1:A3)", &mut r), FormulaValue::Number(60.0));
        assert_eq!(ev("=AVERAGE(A1:A3)", &mut r), FormulaValue::Number(20.0));
        assert_eq!(ev("=MAX(A1:A3)", &mut r), FormulaValue::Number(30.0));
        assert_eq!(ev("=MIN(A1:A3)", &mut r), FormulaValue::Number(10.0));
        assert_eq!(ev("=COUNT(A1:A3)", &mut r), FormulaValue::Number(3.0));
    }

    #[test]
    fn absolute_refs() {
        let mut r = MapResolver::new();
        r.set("B2", FormulaValue::Number(5.0));
        assert_eq!(ev("=$B$2*2", &mut r), FormulaValue::Number(10.0));
    }

    #[test]
    fn strings_and_concat() {
        let mut r = MapResolver::new();
        assert_eq!(
            ev("=\"Hello \"&\"World\"", &mut r),
            FormulaValue::Text("Hello World".into())
        );
        assert_eq!(
            ev("=UPPER(\"abc\")", &mut r),
            FormulaValue::Text("ABC".into())
        );
        assert_eq!(ev("=LEN(\"hello\")", &mut r), FormulaValue::Number(5.0));
        assert_eq!(
            ev("=LEFT(\"hello\",2)", &mut r),
            FormulaValue::Text("he".into())
        );
        assert_eq!(
            ev("=MID(\"hello\",2,3)", &mut r),
            FormulaValue::Text("ell".into())
        );
    }

    #[test]
    fn comparisons_and_if() {
        let mut r = MapResolver::new();
        r.set("A1", FormulaValue::Number(5.0));
        assert_eq!(ev("=A1>3", &mut r), FormulaValue::Bool(true));
        assert_eq!(ev("=A1=5", &mut r), FormulaValue::Bool(true));
        assert_eq!(ev("=A1<>5", &mut r), FormulaValue::Bool(false));
        assert_eq!(
            ev("=IF(A1>3,\"big\",\"small\")", &mut r),
            FormulaValue::Text("big".into())
        );
        assert_eq!(ev("=IF(A1>10,1,0)", &mut r), FormulaValue::Number(0.0));
    }

    #[test]
    fn logical_functions() {
        let mut r = MapResolver::new();
        assert_eq!(ev("=AND(TRUE,TRUE,1)", &mut r), FormulaValue::Bool(true));
        assert_eq!(ev("=AND(TRUE,FALSE)", &mut r), FormulaValue::Bool(false));
        assert_eq!(ev("=OR(FALSE,0,1)", &mut r), FormulaValue::Bool(true));
        assert_eq!(ev("=NOT(FALSE)", &mut r), FormulaValue::Bool(true));
    }

    #[test]
    fn math_functions() {
        let mut r = MapResolver::new();
        assert_eq!(ev("=ABS(-3)", &mut r), FormulaValue::Number(3.0));
        assert_eq!(ev("=SQRT(16)", &mut r), FormulaValue::Number(4.0));
        assert_eq!(ev("=ROUND(1.23456,2)", &mut r), FormulaValue::Number(1.23));
        assert_eq!(ev("=INT(3.9)", &mut r), FormulaValue::Number(3.0));
        assert_eq!(ev("=MOD(7,3)", &mut r), FormulaValue::Number(1.0));
        assert_eq!(ev("=POWER(2,10)", &mut r), FormulaValue::Number(1024.0));
    }

    #[test]
    fn sumif_countif() {
        let mut r = MapResolver::new();
        r.set("A1", FormulaValue::Number(5.0))
            .set("A2", FormulaValue::Number(15.0))
            .set("A3", FormulaValue::Number(25.0));
        assert_eq!(
            ev("=SUMIF(A1:A3,\">10\")", &mut r),
            FormulaValue::Number(40.0)
        );
        assert_eq!(
            ev("=COUNTIF(A1:A3,\">10\")", &mut r),
            FormulaValue::Number(2.0)
        );
    }

    #[test]
    fn sheet_qualified_reference() {
        let mut r = MapResolver::new();
        r.cells.insert(
            (Some("Sheet2".to_string()), 1, 1),
            FormulaValue::Number(42.0),
        );
        assert_eq!(ev("=Sheet2!A1", &mut r), FormulaValue::Number(42.0));
    }

    #[test]
    fn errors_do_not_panic() {
        let mut r = MapResolver::new();
        assert_eq!(
            ev("=BOGUS(1)", &mut r),
            FormulaValue::Error("#NAME?".into())
        );
        assert_eq!(ev("=1+", &mut r), FormulaValue::Error("#VALUE!".into()));
        assert_eq!(ev("=(1+2", &mut r), FormulaValue::Error("#VALUE!".into()));
        assert!(ev("=\"a\"+1", &mut r).is_error());
    }
}
