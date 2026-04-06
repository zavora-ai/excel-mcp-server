use std::sync::Arc;
use tokio::sync::RwLock;

use excel_mcp_server::server::ExcelMcpServer;
use excel_mcp_server::store::WorkbookStore;

#[tokio::main]
async fn main() {
    // Direct all tracing output to stderr so stdout stays clean for MCP protocol.
    tracing_subscriber::fmt()
        .with_writer(std::io::stderr)
        .with_ansi(false)
        .init();

    if let Err(e) = run().await {
        tracing::error!("Fatal error: {:?}", e);
        std::process::exit(1);
    }
}

async fn run() -> anyhow::Result<()> {
    use rmcp::{ServiceExt, transport::stdio};

    let store = Arc::new(RwLock::new(WorkbookStore::new()));
    let server = ExcelMcpServer::new(store);

    let service = server
        .serve(stdio())
        .await
        .inspect_err(|e| tracing::error!("Failed to start MCP service: {:?}", e))?;

    tracing::info!("Excel MCP Server running on stdio");

    service.waiting().await?;
    Ok(())
}
