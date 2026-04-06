use std::sync::Arc;
use tokio::sync::RwLock;

use excel_mcp_server::server::ExcelMcpServer;
use excel_mcp_server::store::WorkbookStore;

#[tokio::main]
async fn main() {
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
    let transport = std::env::args().nth(1).unwrap_or_default();

    match transport.as_str() {
        "http" | "--http" => run_http().await,
        _ => run_stdio().await,
    }
}

/// Stdio transport — for CLI-based MCP clients (Claude Desktop, Cursor, Kiro, etc.)
async fn run_stdio() -> anyhow::Result<()> {
    use rmcp::{ServiceExt, transport::stdio};

    tracing::info!("Excel MCP Server starting on stdio");

    let store = Arc::new(RwLock::new(WorkbookStore::new()));
    let server = ExcelMcpServer::new(store);

    let service = server
        .serve(stdio())
        .await
        .inspect_err(|e| tracing::error!("Failed to start MCP service: {:?}", e))?;

    service.waiting().await?;
    Ok(())
}

/// Streamable HTTP transport — for web-based MCP clients
async fn run_http() -> anyhow::Result<()> {
    use rmcp::transport::streamable_http_server::{
        StreamableHttpServerConfig, StreamableHttpService,
        session::local::LocalSessionManager,
    };

    let bind = std::env::var("BIND_ADDRESS").unwrap_or_else(|_| "127.0.0.1:8080".to_string());
    let ct = tokio_util::sync::CancellationToken::new();

    let service = StreamableHttpService::new(
        || {
            let store = Arc::new(RwLock::new(WorkbookStore::new()));
            Ok(ExcelMcpServer::new(store))
        },
        LocalSessionManager::default().into(),
        StreamableHttpServerConfig::default().with_cancellation_token(ct.child_token()),
    );

    let router = axum::Router::new().nest_service("/mcp", service);
    let listener = tokio::net::TcpListener::bind(&bind).await?;

    tracing::info!("Excel MCP Server listening on http://{}/mcp", bind);

    axum::serve(listener, router)
        .with_graceful_shutdown(async move {
            tokio::signal::ctrl_c().await.unwrap();
            tracing::info!("Shutting down...");
            ct.cancel();
        })
        .await?;

    Ok(())
}
