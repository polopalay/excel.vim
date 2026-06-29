//! Bộ tính công thức Excel tối giản, viết tay (không dùng crate parser
//! ngoài, đúng tinh thần của toàn bộ project: "viết tay bằng quick-xml để
//! tránh phụ thuộc 1 crate ngoài có API có thể thay đổi") — đủ dùng cho
//! các công thức phổ biến: số học, so sánh, nối chuỗi, tham chiếu
//! cell/range, và 1 bộ hàm cơ bản (SUM, IF, AVERAGE, ...).
//!
//! Đánh giá LƯỜI + GHI NHỚ (lazy + memoized): mỗi cell formula được tính
//! qua `eval_cell`, có cache theo (row,col) và 1 tập `visiting` để phát
//! hiện tham chiếu vòng (circular reference) -> trả "#CIRCULAR!" thay vì
//! đệ quy vô hạn / stack overflow.
//!
//! GIỚI HẠN (xem README phần "Giới hạn" để biết chi tiết khi mở rộng sau):
//! - Không hỗ trợ shared formula (<f t="shared">), tham chiếu sheet khác
//!   (Sheet2!A1), hay named range.
//! - Bộ hàm còn ít: SUM, AVERAGE, MIN, MAX, COUNT, ABS, ROUND, IF, AND, OR,
//!   NOT, CONCAT/CONCATENATE. Thêm hàm mới chỉ cần sửa `eval_call`.
//! - Không áp number format (currency/percent/date) — luôn ra số thô.

use std::collections::{HashMap, HashSet};

use crate::coord::ref_to_col_row;
use crate::model::SheetData;

// ----------------------------------------------------------------------------
// Value: kết quả trung gian / cuối cùng của 1 phép tính
// ----------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Number(f64),
    Text(String),
    Bool(bool),
    /// Lỗi dạng Excel, ví dụ "#DIV/0!", "#REF!", "#VALUE!", "#NAME?",
    /// "#CIRCULAR!" (không phải mã lỗi chuẩn Excel, nhưng đủ rõ nghĩa).
    Error(String),
    Empty,
}

impl Value {
    fn as_number(&self) -> Result<f64, Value> {
        match self {
            Value::Number(n) => Ok(*n),
            Value::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
            Value::Empty => Ok(0.0),
            Value::Text(s) => s
                .trim()
                .parse::<f64>()
                .map_err(|_| Value::Error("#VALUE!".to_string())),
            Value::Error(e) => Err(Value::Error(e.clone())),
        }
    }

    /// Chuyển Value thành text để ghi vào sheet.cells (giống cách Excel
    /// cache kết quả công thức vào <v>).
    pub fn to_display_string(&self) -> String {
        match self {
            Value::Number(n) => format_number(*n),
            Value::Text(s) => s.clone(),
            Value::Bool(b) => {
                if *b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            Value::Error(e) => e.clone(),
            Value::Empty => String::new(),
        }
    }
}

/// Format số kiểu Excel: số nguyên -> không có ".0", số thực -> cắt bớt
/// sai số dấu phẩy động (0.1+0.2 = 0.30000000000000004) để hiển thị gọn.
fn format_number(n: f64) -> String {
    if !n.is_finite() {
        return "#NUM!".to_string();
    }
    if n == n.trunc() && n.abs() < 1e15 {
        return format!("{}", n as i64);
    }
    let mut s = format!("{n:.10}");
    while s.ends_with('0') {
        s.pop();
    }
    if s.ends_with('.') {
        s.pop();
    }
    s
}

fn truthy(v: &Value) -> bool {
    match v {
        Value::Bool(b) => *b,
        Value::Number(n) => *n != 0.0,
        Value::Text(s) => !s.is_empty() && s.to_uppercase() != "FALSE",
        Value::Empty => false,
        Value::Error(_) => false,
    }
}

/// So sánh 2 Value: nếu cả 2 quy đổi được sang số -> so sánh số học,
/// ngược lại so sánh text (không phân biệt hoa/thường, giống Excel).
fn compare_values(l: &Value, r: &Value) -> std::cmp::Ordering {
    match (l.as_number(), r.as_number()) {
        (Ok(a), Ok(b)) => a.partial_cmp(&b).unwrap_or(std::cmp::Ordering::Equal),
        _ => l
            .to_display_string()
            .to_uppercase()
            .cmp(&r.to_display_string().to_uppercase()),
    }
}

// ----------------------------------------------------------------------------
// Tokenizer
// ----------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq)]
enum Tok {
    Num(f64),
    Str(String),
    Ref(String),   // "A1" (đã bỏ dấu $ tuyệt đối)
    Ident(String), // tên hàm (SUM...) hoặc literal TRUE/FALSE
    LParen,
    RParen,
    Comma,
    Colon,
    Plus,
    Minus,
    Star,
    Slash,
    Caret,
    Amp,
    Percent,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}

/// Tách 1 chuỗi formula (KHÔNG có dấu "=" đầu) thành danh sách token.
/// Cell ref được nhận diện bằng pattern "chữ + số liền nhau" (ví dụ A1,
/// $B$12) ngay trong lúc quét ký tự, không cần token IDENT riêng rồi ghép
/// lại — tránh nhầm với tên hàm có số trong tên (LOG10...).
fn tokenize(src: &str) -> Result<Vec<Tok>, Value> {
    let chars: Vec<char> = src.chars().collect();
    let mut i = 0;
    let mut out = Vec::new();
    while i < chars.len() {
        let c = chars[i];
        match c {
            ' ' | '\t' => i += 1,
            '(' => {
                out.push(Tok::LParen);
                i += 1;
            }
            ')' => {
                out.push(Tok::RParen);
                i += 1;
            }
            ',' => {
                out.push(Tok::Comma);
                i += 1;
            }
            ':' => {
                out.push(Tok::Colon);
                i += 1;
            }
            '+' => {
                out.push(Tok::Plus);
                i += 1;
            }
            '-' => {
                out.push(Tok::Minus);
                i += 1;
            }
            '*' => {
                out.push(Tok::Star);
                i += 1;
            }
            '/' => {
                out.push(Tok::Slash);
                i += 1;
            }
            '^' => {
                out.push(Tok::Caret);
                i += 1;
            }
            '&' => {
                out.push(Tok::Amp);
                i += 1;
            }
            '%' => {
                out.push(Tok::Percent);
                i += 1;
            }
            '=' => {
                out.push(Tok::Eq);
                i += 1;
            }
            '<' => {
                if chars.get(i + 1) == Some(&'>') {
                    out.push(Tok::Ne);
                    i += 2;
                } else if chars.get(i + 1) == Some(&'=') {
                    out.push(Tok::Le);
                    i += 2;
                } else {
                    out.push(Tok::Lt);
                    i += 1;
                }
            }
            '>' => {
                if chars.get(i + 1) == Some(&'=') {
                    out.push(Tok::Ge);
                    i += 2;
                } else {
                    out.push(Tok::Gt);
                    i += 1;
                }
            }
            '"' => {
                let mut s = String::new();
                i += 1;
                while i < chars.len() && chars[i] != '"' {
                    s.push(chars[i]);
                    i += 1;
                }
                i += 1; // bỏ qua dấu " đóng (nếu chuỗi không kết thúc đúng, chấp nhận cắt)
                out.push(Tok::Str(s));
            }
            '0'..='9' | '.' => {
                let start = i;
                while i < chars.len() && (chars[i].is_ascii_digit() || chars[i] == '.') {
                    i += 1;
                }
                // hỗ trợ ký hiệu khoa học: 1e10, 2.5E-3
                if i < chars.len() && (chars[i] == 'e' || chars[i] == 'E') {
                    let mut j = i + 1;
                    if j < chars.len() && (chars[j] == '+' || chars[j] == '-') {
                        j += 1;
                    }
                    if j < chars.len() && chars[j].is_ascii_digit() {
                        while j < chars.len() && chars[j].is_ascii_digit() {
                            j += 1;
                        }
                        i = j;
                    }
                }
                let num_str: String = chars[start..i].iter().collect();
                let n = num_str
                    .parse::<f64>()
                    .map_err(|_| Value::Error("#VALUE!".to_string()))?;
                out.push(Tok::Num(n));
            }
            '$' | 'a'..='z' | 'A'..='Z' => {
                let start = i;
                while i < chars.len() && (chars[i].is_ascii_alphabetic() || chars[i] == '$') {
                    i += 1;
                }
                let letters_part: String = chars[start..i].iter().collect();
                if i < chars.len() && chars[i].is_ascii_digit() {
                    // chữ + số liền nhau ngay sau -> cell ref (A1, $A$1, ...)
                    let digit_start = i;
                    while i < chars.len() && chars[i].is_ascii_digit() {
                        i += 1;
                    }
                    let digits: String = chars[digit_start..i].iter().collect();
                    let clean_ref = format!("{}{}", letters_part.replace('$', ""), digits);
                    out.push(Tok::Ref(clean_ref.to_uppercase()));
                } else {
                    out.push(Tok::Ident(letters_part.to_uppercase()));
                }
            }
            _ => return Err(Value::Error("#NAME?".to_string())),
        }
    }
    Ok(out)
}

// ----------------------------------------------------------------------------
// Parser (recursive descent) -> AST
// ----------------------------------------------------------------------------

#[derive(Debug, Clone)]
enum Expr {
    Num(f64),
    Str(String),
    Bool(bool),
    CellRef(String),
    RangeRef(String, String),
    Neg(Box<Expr>),
    BinOp(Tok, Box<Expr>, Box<Expr>),
    Call(String, Vec<Expr>),
}

struct Parser {
    toks: Vec<Tok>,
    pos: usize,
}

impl Parser {
    fn peek(&self) -> Option<&Tok> {
        self.toks.get(self.pos)
    }

    fn next(&mut self) -> Option<Tok> {
        let t = self.toks.get(self.pos).cloned();
        self.pos += 1;
        t
    }

    // Thứ tự ưu tiên (thấp -> cao):
    //   so sánh (=,<>,<,<=,>,>=) < nối chuỗi (&) < +- < */ < ^ < unary -/+ < %
    fn parse_expr(&mut self) -> Result<Expr, Value> {
        let mut left = self.parse_concat()?;
        while matches!(
            self.peek(),
            Some(Tok::Eq | Tok::Ne | Tok::Lt | Tok::Le | Tok::Gt | Tok::Ge)
        ) {
            let op = self.next().unwrap();
            let right = self.parse_concat()?;
            left = Expr::BinOp(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn parse_concat(&mut self) -> Result<Expr, Value> {
        let mut left = self.parse_additive()?;
        while matches!(self.peek(), Some(Tok::Amp)) {
            self.next();
            let right = self.parse_additive()?;
            left = Expr::BinOp(Tok::Amp, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn parse_additive(&mut self) -> Result<Expr, Value> {
        let mut left = self.parse_mult()?;
        while matches!(self.peek(), Some(Tok::Plus | Tok::Minus)) {
            let op = self.next().unwrap();
            let right = self.parse_mult()?;
            left = Expr::BinOp(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn parse_mult(&mut self) -> Result<Expr, Value> {
        let mut left = self.parse_power()?;
        while matches!(self.peek(), Some(Tok::Star | Tok::Slash)) {
            let op = self.next().unwrap();
            let right = self.parse_power()?;
            left = Expr::BinOp(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn parse_power(&mut self) -> Result<Expr, Value> {
        let base = self.parse_unary()?;
        if matches!(self.peek(), Some(Tok::Caret)) {
            self.next();
            let exp = self.parse_power()?; // ^ kết hợp phải (right-associative)
            return Ok(Expr::BinOp(Tok::Caret, Box::new(base), Box::new(exp)));
        }
        Ok(base)
    }

    fn parse_unary(&mut self) -> Result<Expr, Value> {
        if matches!(self.peek(), Some(Tok::Minus)) {
            self.next();
            return Ok(Expr::Neg(Box::new(self.parse_unary()?)));
        }
        if matches!(self.peek(), Some(Tok::Plus)) {
            self.next();
            return self.parse_unary();
        }
        self.parse_postfix()
    }

    /// Hậu tố "%" (ví dụ 50% = 0.5) — quy về phép chia /100.
    fn parse_postfix(&mut self) -> Result<Expr, Value> {
        let mut e = self.parse_primary()?;
        while matches!(self.peek(), Some(Tok::Percent)) {
            self.next();
            e = Expr::BinOp(Tok::Slash, Box::new(e), Box::new(Expr::Num(100.0)));
        }
        Ok(e)
    }

    fn parse_primary(&mut self) -> Result<Expr, Value> {
        match self.next() {
            Some(Tok::Num(n)) => Ok(Expr::Num(n)),
            Some(Tok::Str(s)) => Ok(Expr::Str(s)),
            Some(Tok::Ref(r)) => {
                if matches!(self.peek(), Some(Tok::Colon)) {
                    self.next();
                    match self.next() {
                        Some(Tok::Ref(r2)) => Ok(Expr::RangeRef(r, r2)),
                        _ => Err(Value::Error("#REF!".to_string())),
                    }
                } else {
                    Ok(Expr::CellRef(r))
                }
            }
            Some(Tok::Ident(name)) => {
                if matches!(self.peek(), Some(Tok::LParen)) {
                    self.next();
                    let mut args = Vec::new();
                    if !matches!(self.peek(), Some(Tok::RParen)) {
                        loop {
                            args.push(self.parse_expr()?);
                            if matches!(self.peek(), Some(Tok::Comma)) {
                                self.next();
                            } else {
                                break;
                            }
                        }
                    }
                    if !matches!(self.next(), Some(Tok::RParen)) {
                        return Err(Value::Error("#NAME?".to_string()));
                    }
                    Ok(Expr::Call(name, args))
                } else if name == "TRUE" {
                    Ok(Expr::Bool(true))
                } else if name == "FALSE" {
                    Ok(Expr::Bool(false))
                } else {
                    // Identifier trơn không phải hàm/TRUE/FALSE -> coi như
                    // named range, CHƯA hỗ trợ.
                    Err(Value::Error("#NAME?".to_string()))
                }
            }
            Some(Tok::LParen) => {
                let e = self.parse_expr()?;
                if !matches!(self.next(), Some(Tok::RParen)) {
                    return Err(Value::Error("#VALUE!".to_string()));
                }
                Ok(e)
            }
            _ => Err(Value::Error("#VALUE!".to_string())),
        }
    }
}

// ----------------------------------------------------------------------------
// Evaluator
// ----------------------------------------------------------------------------

type Cache = HashMap<(u32, u32), Value>;
type Visiting = HashSet<(u32, u32)>;

/// Diễn giải nội dung 1 cell KHÔNG có formula (literal) thành Value, dùng
/// đúng quy ước hiển thị hiện có của plugin: "" -> rỗng, parse được số ->
/// số, "TRUE"/"FALSE" -> bool, còn lại -> text.
fn parse_literal(s: &str) -> Value {
    if s.is_empty() {
        Value::Empty
    } else if let Ok(n) = s.parse::<f64>() {
        Value::Number(n)
    } else if s == "TRUE" {
        Value::Bool(true)
    } else if s == "FALSE" {
        Value::Bool(false)
    } else {
        Value::Text(s.to_string())
    }
}

/// Tokenize + parse + eval 1 chuỗi formula (không có dấu "=" đầu) thành
/// Value. Tách riêng khỏi eval_cell() để luồng mượn `cache`/`visiting` rõ
/// ràng, dễ đọc hơn là nhúng 1 closure gọi ngay (IIFE) bên trong eval_cell.
fn eval_formula_text(
    formula: &str,
    sheet: &SheetData,
    cache: &mut Cache,
    visiting: &mut Visiting,
) -> Value {
    let toks = match tokenize(formula) {
        Ok(t) => t,
        Err(e) => return e,
    };
    let mut parser = Parser { toks, pos: 0 };
    match parser.parse_expr() {
        Ok(ast) => eval_expr(&ast, sheet, cache, visiting),
        Err(e) => e,
    }
}

/// Tính giá trị của 1 cell (row, col), có cache + phát hiện vòng lặp.
/// Nếu cell có formula trong `sheet.formulas` -> parse + eval công thức đó
/// (đệ quy gọi lại eval_cell cho mọi cell mà công thức tham chiếu tới).
/// Nếu không có formula -> dùng nội dung thô hiện tại trong `sheet.cells`.
fn eval_cell(
    sheet: &SheetData,
    row: u32,
    col: u32,
    cache: &mut Cache,
    visiting: &mut Visiting,
) -> Value {
    if let Some(v) = cache.get(&(row, col)) {
        return v.clone();
    }
    if visiting.contains(&(row, col)) {
        return Value::Error("#CIRCULAR!".to_string());
    }

    let value = if let Some(formula) = sheet.formulas.get(&(row, col)).cloned() {
        visiting.insert((row, col));
        let result = eval_formula_text(&formula, sheet, cache, visiting);
        visiting.remove(&(row, col));
        result
    } else {
        parse_literal(sheet.get(row, col).unwrap_or(""))
    };

    cache.insert((row, col), value.clone());
    value
}

fn eval_expr(e: &Expr, sheet: &SheetData, cache: &mut Cache, visiting: &mut Visiting) -> Value {
    match e {
        Expr::Num(n) => Value::Number(*n),
        Expr::Str(s) => Value::Text(s.clone()),
        Expr::Bool(b) => Value::Bool(*b),
        Expr::Neg(inner) => match eval_expr(inner, sheet, cache, visiting).as_number() {
            Ok(n) => Value::Number(-n),
            Err(e) => e,
        },
        Expr::CellRef(r) => match ref_to_col_row(r) {
            Ok((c, rr)) => eval_cell(sheet, rr, c, cache, visiting),
            Err(_) => Value::Error("#REF!".to_string()),
        },
        // Range đứng riêng (không nằm trong hàm tổng hợp) không có nghĩa
        // là 1 giá trị đơn -> giống Excel trả "#VALUE!".
        Expr::RangeRef(..) => Value::Error("#VALUE!".to_string()),
        Expr::BinOp(op, l, r) => {
            let lv = eval_expr(l, sheet, cache, visiting);
            if let Value::Error(e) = &lv {
                return Value::Error(e.clone());
            }
            let rv = eval_expr(r, sheet, cache, visiting);
            if let Value::Error(e) = &rv {
                return Value::Error(e.clone());
            }
            apply_binop(op, lv, rv)
        }
        Expr::Call(name, args) => eval_call(name, args, sheet, cache, visiting),
    }
}

fn apply_binop(op: &Tok, l: Value, r: Value) -> Value {
    match op {
        Tok::Amp => Value::Text(format!("{}{}", l.to_display_string(), r.to_display_string())),
        Tok::Plus | Tok::Minus | Tok::Star | Tok::Slash | Tok::Caret => {
            let (ln, rn) = match (l.as_number(), r.as_number()) {
                (Ok(a), Ok(b)) => (a, b),
                (Err(e), _) | (_, Err(e)) => return e,
            };
            match op {
                Tok::Plus => Value::Number(ln + rn),
                Tok::Minus => Value::Number(ln - rn),
                Tok::Star => Value::Number(ln * rn),
                Tok::Slash => {
                    if rn == 0.0 {
                        Value::Error("#DIV/0!".to_string())
                    } else {
                        Value::Number(ln / rn)
                    }
                }
                Tok::Caret => Value::Number(ln.powf(rn)),
                _ => unreachable!(),
            }
        }
        Tok::Eq | Tok::Ne | Tok::Lt | Tok::Le | Tok::Gt | Tok::Ge => {
            let cmp = compare_values(&l, &r);
            Value::Bool(match op {
                Tok::Eq => cmp == std::cmp::Ordering::Equal,
                Tok::Ne => cmp != std::cmp::Ordering::Equal,
                Tok::Lt => cmp == std::cmp::Ordering::Less,
                Tok::Le => cmp != std::cmp::Ordering::Greater,
                Tok::Gt => cmp == std::cmp::Ordering::Greater,
                Tok::Ge => cmp != std::cmp::Ordering::Less,
                _ => unreachable!(),
            })
        }
        _ => Value::Error("#VALUE!".to_string()),
    }
}

/// "Trải" 1 tham số hàm thành danh sách Value: range (A1:A5) -> nhiều giá
/// trị (mỗi cell 1 phần tử), biểu thức khác -> đúng 1 giá trị. Cho phép
/// SUM(A1:A5, B2, 10) kiểu Excel hoạt động đúng.
fn collect_values(
    e: &Expr,
    sheet: &SheetData,
    cache: &mut Cache,
    visiting: &mut Visiting,
) -> Vec<Value> {
    match e {
        Expr::RangeRef(a, b) => {
            let (c1, r1) = ref_to_col_row(a).unwrap_or((0, 0));
            let (c2, r2) = ref_to_col_row(b).unwrap_or((0, 0));
            let mut out = Vec::new();
            for rr in r1.min(r2)..=r1.max(r2) {
                for cc in c1.min(c2)..=c1.max(c2) {
                    out.push(eval_cell(sheet, rr, cc, cache, visiting));
                }
            }
            out
        }
        other => vec![eval_expr(other, sheet, cache, visiting)],
    }
}

/// SUM bỏ qua text/rỗng trong range (giống Excel), nhưng vẫn lan truyền lỗi
/// thật (#REF!, #DIV/0!...) nếu gặp trong range.
fn sum_values(values: &[Value]) -> Value {
    let mut acc = 0.0;
    for v in values {
        match v {
            Value::Error(e) => return Value::Error(e.clone()),
            Value::Number(n) => acc += n,
            Value::Bool(b) => acc += if *b { 1.0 } else { 0.0 },
            Value::Text(_) | Value::Empty => {}
        }
    }
    Value::Number(acc)
}

fn fold_minmax(values: &[Value], f: impl Fn(f64, f64) -> f64) -> Value {
    let nums: Vec<f64> = values
        .iter()
        .filter_map(|v| match v {
            Value::Number(n) => Some(*n),
            _ => None,
        })
        .collect();
    if nums.is_empty() {
        return Value::Number(0.0);
    }
    Value::Number(nums[1..].iter().fold(nums[0], |acc, &n| f(acc, n)))
}

fn eval_call(
    name: &str,
    args: &[Expr],
    sheet: &SheetData,
    cache: &mut Cache,
    visiting: &mut Visiting,
) -> Value {
    // IF cần short-circuit (chỉ eval nhánh được chọn) -> xử lý riêng, KHÔNG
    // đi qua bước "trải toàn bộ tham số" ở dưới (range trong IF không có
    // ý nghĩa "trải", và eval cả 2 nhánh là sai semantics + tốn công vô ích
    // / có thể gây lỗi không đáng có ở nhánh không được chọn).
    if name == "IF" {
        if args.is_empty() {
            return Value::Error("#VALUE!".to_string());
        }
        let cond = eval_expr(&args[0], sheet, cache, visiting);
        if let Value::Error(e) = &cond {
            return Value::Error(e.clone());
        }
        return if truthy(&cond) {
            args.get(1)
                .map(|e| eval_expr(e, sheet, cache, visiting))
                .unwrap_or(Value::Bool(true))
        } else {
            args.get(2)
                .map(|e| eval_expr(e, sheet, cache, visiting))
                .unwrap_or(Value::Bool(false))
        };
    }

    let flat: Vec<Value> = args
        .iter()
        .flat_map(|a| collect_values(a, sheet, cache, visiting))
        .collect();

    match name {
        "SUM" => sum_values(&flat),
        "AVERAGE" => {
            let nums: Vec<f64> = flat
                .iter()
                .filter_map(|v| match v {
                    Value::Number(n) => Some(*n),
                    Value::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
                    _ => None,
                })
                .collect();
            if nums.is_empty() {
                Value::Error("#DIV/0!".to_string())
            } else {
                Value::Number(nums.iter().sum::<f64>() / nums.len() as f64)
            }
        }
        "MIN" => fold_minmax(&flat, f64::min),
        "MAX" => fold_minmax(&flat, f64::max),
        "COUNT" => Value::Number(flat.iter().filter(|v| matches!(v, Value::Number(_))).count() as f64),
        "ABS" => match flat.first().map(Value::as_number) {
            Some(Ok(n)) => Value::Number(n.abs()),
            Some(Err(e)) => e,
            None => Value::Error("#VALUE!".to_string()),
        },
        "ROUND" => {
            let n = flat.first().and_then(|v| v.as_number().ok()).unwrap_or(0.0);
            let digits = flat.get(1).and_then(|v| v.as_number().ok()).unwrap_or(0.0) as i32;
            let factor = 10f64.powi(digits);
            Value::Number((n * factor).round() / factor)
        }
        "AND" => Value::Bool(flat.iter().all(truthy)),
        "OR" => Value::Bool(flat.iter().any(truthy)),
        "NOT" => Value::Bool(!truthy(flat.first().unwrap_or(&Value::Bool(false)))),
        "CONCAT" | "CONCATENATE" => {
            Value::Text(flat.iter().map(Value::to_display_string).collect::<Vec<_>>().concat())
        }
        _ => Value::Error("#NAME?".to_string()),
    }
}

/// Tính lại TOÀN BỘ formula trong sheet — gọi sau khi áp dụng nội dung
/// bảng ASCII mới (save_logic::apply_ascii_table), TRƯỚC KHI ghi ra XML.
/// Ghi kết quả hiển thị vào `sheet.cells`, đúng như cách Excel cache giá
/// trị <v> song song với <f>.
pub fn evaluate_all(sheet: &mut SheetData) {
    let mut cache: Cache = HashMap::new();
    let coords: Vec<(u32, u32)> = sheet.formulas.keys().copied().collect();
    for &(row, col) in &coords {
        let mut visiting: Visiting = HashSet::new();
        eval_cell(sheet, row, col, &mut cache, &mut visiting);
    }
    for &(row, col) in &coords {
        if let Some(v) = cache.get(&(row, col)) {
            sheet.set(row, col, v.to_display_string());
        }
    }
}

// ----------------------------------------------------------------------------
// Show Formulas (:ExcelShowFormula) và Apply Formula (:ExcelApplyFormula)
// ----------------------------------------------------------------------------

/// Tạo 1 bản sao SheetData mà các cell CÓ formula sẽ hiển thị TEXT công
/// thức (kèm dấu "=" đầu) thay vì giá trị đã tính — dùng cho mode "hiện
/// công thức" (:ExcelShowFormula, giống Ctrl+` trong Excel). Các cell
/// khác giữ nguyên y nguyên. Merge/style/formulas KHÔNG đổi — chỉ đổi
/// `cells` để hiển thị, nên display.rs/table.rs (đo độ rộng cột theo nội
/// dung) tự động hoạt động đúng mà không cần sửa gì thêm.
pub fn sheet_with_formula_text_shown(sheet: &SheetData) -> SheetData {
    let mut shown = sheet.clone();
    for (&(row, col), text) in &sheet.formulas {
        shown.set(row, col, format!("={text}"));
    }
    shown
}

/// Dịch (shift) mọi cell reference trong 1 chuỗi formula theo
/// (row_delta, col_delta) — dùng cho lệnh "Apply Formula" (giống kéo fill
/// handle trong Excel): ô mẫu B2 = "B1+1", kéo xuống B3 sẽ tự thành
/// "B2+1" (row_delta=+1), kéo sang C2 sẽ thành "C1+1" (col_delta=+1).
///
/// QUY TẮC tương đối/tuyệt đối kiểu Excel: phần KHÔNG có dấu "$" đứng
/// ngay trước sẽ bị dịch theo delta, phần CÓ "$" giữ nguyên — ví dụ
/// "$A1" kéo ngang vẫn giữ cột A (có "$"), nhưng kéo xuống vẫn dịch dòng
/// (không có "$" trước số dòng).
///
/// Chuỗi trong dấu nháy kép ("...") được giữ nguyên, không coi text bên
/// trong là cell reference. Tên hàm có chứa số (ví dụ giả định "LOG10")
/// cũng được giữ nguyên — nhận biết bằng cách: nếu ngay sau phần
/// "chữ+số" là dấu "(" thì đó là tên hàm, không phải cell ref.
///
/// GIỚI HẠN: không tham chiếu sheet khác/named range (đã là giới hạn có
/// từ trước của module này), và không tự phát hiện #REF! khi shift ra
/// ngoài biên (row/col < 1 sẽ bị kẹp về 1 thay vì báo lỗi).
pub fn shift_formula_refs(formula: &str, row_delta: i64, col_delta: i64) -> String {
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len());
    let mut i = 0;
    let mut in_string = false;

    while i < chars.len() {
        let c = chars[i];

        if in_string {
            out.push(c);
            if c == '"' {
                in_string = false;
            }
            i += 1;
            continue;
        }
        if c == '"' {
            in_string = true;
            out.push(c);
            i += 1;
            continue;
        }

        if c == '$' || c.is_ascii_alphabetic() {
            let start = i;
            let mut j = i;
            // phần chữ (+ dấu $ tuỳ chọn đứng trước chữ, ví dụ "$A")
            while j < chars.len() && (chars[j].is_ascii_alphabetic() || chars[j] == '$') {
                j += 1;
            }
            let letters_raw = &chars[start..j];
            let col_absolute = letters_raw.first() == Some(&'$');
            let letters: String = letters_raw.iter().filter(|c| **c != '$').collect();

            let is_candidate = !letters.is_empty()
                && j < chars.len()
                && (chars[j].is_ascii_digit() || chars[j] == '$');

            if is_candidate {
                let row_absolute = chars[j] == '$';
                let digit_start = if row_absolute { j + 1 } else { j };
                let mut k = digit_start;
                while k < chars.len() && chars[k].is_ascii_digit() {
                    k += 1;
                }

                if k > digit_start {
                    // Nếu ngay sau là "(" -> đây là TÊN HÀM kiểu "LOG10(",
                    // không phải cell reference -> giữ nguyên, không dịch.
                    if k < chars.len() && chars[k] == '(' {
                        out.extend(chars[start..k].iter());
                        i = k;
                        continue;
                    }

                    let digits: String = chars[digit_start..k].iter().collect();
                    let old_col = crate::coord::letters_to_col(&letters);
                    let old_row: i64 = digits.parse().unwrap_or(1);

                    let new_col = if col_absolute {
                        old_col
                    } else {
                        (old_col as i64 + col_delta).max(1) as u32
                    };
                    let new_row = if row_absolute {
                        old_row
                    } else {
                        (old_row + row_delta).max(1)
                    };

                    if col_absolute {
                        out.push('$');
                    }
                    out.push_str(&crate::coord::col_to_letters(new_col));
                    if row_absolute {
                        out.push('$');
                    }
                    out.push_str(&new_row.to_string());

                    i = k;
                    continue;
                }
            }

            // Không phải cell ref hợp lệ (tên hàm SUM/IF/..., hoặc "$" lẻ)
            // -> giữ nguyên y nguyên đoạn vừa quét.
            out.extend(letters_raw.iter());
            i = j;
            continue;
        }

        out.push(c);
        i += 1;
    }

    out
}
