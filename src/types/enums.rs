use schemars::JsonSchema;
use serde::Deserialize;

#[derive(Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum WaterfallPointKind {
    Increase,
    Decrease,
    Total,
}

#[derive(Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum ShapeKind {
    Rectangle,
    RoundedRectangle,
    Ellipse,
    Triangle,
    Diamond,
    Arrow,
    Callout,
    TextBox,
}

#[derive(Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum ChartType {
    Bar,
    Column,
    Line,
    Pie,
    Scatter,
    Area,
    Doughnut,
}

#[derive(Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum SparklineType {
    Line,
    Column,
    WinLoss,
}

#[derive(Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum HorizontalAlignment {
    Left,
    Center,
    Right,
    Fill,
    Justify,
}

#[derive(Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum VerticalAlignment {
    Top,
    Center,
    Bottom,
    Justify,
}

#[derive(Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum BorderStyle {
    Thin,
    Medium,
    Thick,
    Dashed,
    Dotted,
    Double,
    None,
}

#[derive(Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum LegendPosition {
    Top,
    Bottom,
    Left,
    Right,
    None,
}

#[derive(Default, Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum MatchMode {
    Exact,
    #[default]
    #[serde(rename = "substring")]
    Substring,
}

#[derive(Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum ComparisonOperator {
    GreaterThan,
    LessThan,
    Between,
    EqualTo,
    NotEqualTo,
    GreaterThanOrEqual,
    LessThanOrEqual,
}

#[derive(Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum IconSetStyle {
    ThreeArrows,
    ThreeTrafficLights,
    ThreeSymbols,
    FourArrows,
    FiveArrows,
}

#[derive(Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum AlertStyle {
    Stop,
    Warning,
    Information,
}

#[derive(Deserialize, JsonSchema)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum ConditionalFormatRule {
    CellValue {
        operator: ComparisonOperator,
        value: f64,
        value2: Option<f64>,
    },
    ColorScale2 {
        min_color: String,
        max_color: String,
    },
    ColorScale3 {
        min_color: String,
        mid_color: String,
        max_color: String,
    },
    DataBar {
        color: String,
    },
    IconSet {
        style: IconSetStyle,
    },
}

#[derive(Deserialize, JsonSchema)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum ValidationRule {
    /// Dropdown list from explicit values
    List { values: Vec<String> },
    /// Dropdown list from a cell range
    ListRange { range: String },
    /// Whole number in a range
    WholeNumber { min: Option<i64>, max: Option<i64> },
    /// Decimal number in a range
    Decimal { min: Option<f64>, max: Option<f64> },
    /// Date in a range (ISO 8601 strings)
    DateRange {
        min: Option<String>,
        max: Option<String>,
    },
    /// Text length in a range
    TextLength { min: Option<u32>, max: Option<u32> },
    /// Custom formula
    CustomFormula { formula: String },
}

#[derive(Deserialize, JsonSchema)]
pub struct ConditionalFormatStyle {
    #[serde(default)]
    pub font_color: Option<String>,
    #[serde(default)]
    pub background_color: Option<String>,
    #[serde(default)]
    pub bold: Option<bool>,
}

#[derive(Deserialize, JsonSchema)]
pub struct ValidationMessage {
    pub title: String,
    pub body: String,
}

#[derive(Deserialize, JsonSchema)]
pub struct ValidationAlert {
    pub style: AlertStyle,
    pub title: String,
    pub message: String,
}
