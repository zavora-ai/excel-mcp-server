use std::fmt;

/// Central error type for the Excel MCP server.
///
/// Each variant carries a descriptive `String` message so that callers
/// (ultimately the LLM client) receive actionable feedback.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ExcelMcpError {
    /// The requested resource (workbook, sheet, cell, etc.) was not found.
    NotFound(String),
    /// The caller supplied an invalid parameter value.
    InvalidInput(String),
    /// The requested operation is not supported by the engine that owns the workbook.
    EngineUnsupported(String),
    /// The workbook store has reached its maximum capacity.
    CapacityExceeded(String),
    /// An I/O error occurred (file read/write, path issues, etc.).
    IoError(String),
    /// A parsing error occurred (cell references, data formats, etc.).
    ParseError(String),
    /// The workbook was evicted from the store due to inactivity.
    Evicted(String),
}

impl fmt::Display for ExcelMcpError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            ExcelMcpError::NotFound(msg) => write!(f, "Not found: {msg}"),
            ExcelMcpError::InvalidInput(msg) => write!(f, "Invalid input: {msg}"),
            ExcelMcpError::EngineUnsupported(msg) => write!(f, "Engine unsupported: {msg}"),
            ExcelMcpError::CapacityExceeded(msg) => write!(f, "Capacity exceeded: {msg}"),
            ExcelMcpError::IoError(msg) => write!(f, "IO error: {msg}"),
            ExcelMcpError::ParseError(msg) => write!(f, "Parse error: {msg}"),
            ExcelMcpError::Evicted(msg) => write!(f, "Evicted: {msg}"),
        }
    }
}

impl std::error::Error for ExcelMcpError {}

impl From<std::io::Error> for ExcelMcpError {
    fn from(err: std::io::Error) -> Self {
        ExcelMcpError::IoError(err.to_string())
    }
}

impl From<serde_json::Error> for ExcelMcpError {
    fn from(err: serde_json::Error) -> Self {
        ExcelMcpError::ParseError(err.to_string())
    }
}
